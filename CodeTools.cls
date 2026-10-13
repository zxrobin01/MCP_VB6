VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CodeTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' All MCP tools operating on the VB6 IDE (VBIDE extensibility) and the
' project files on disk. Covers the complete VBIDE 5.3 object model:
' VBE / VBProjects / VBProject / VBComponents / VBComponent / CodeModule /
' CodePane / References / Properties / Members / Windows / form designer /
' AddIns / CommandBars. Every tool returns JSON text.

Private Const DEFAULT_SEARCH_PATTERNS As String = "*.bas;*.frm;*.cls;*.vbp;*.ctl;*.pag;*.dsr;*.dob;*.dox;*.inc;*.txt;*.md"

Private m_vbe As VBIDE.VBE
Private m_server As Object
' TASK-P1-101: diagnostics for the LAST FindComponent miss. A component name can contain
' anything (including the | that separates a MCPREASON tag), so the strings live here and the
' tag stays a bare reason; McpReasonJson reads them back when reason = component_not_found.
Private m_cnfTried As String
Private m_cnfCount As Long
Private m_cnfNames As Object
Private m_cnfNear As Object
Private m_cnfTrimmed As String
Private m_cnfCollapsed As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function PostMessageW Lib "user32" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function FindWindowExW Lib "user32" (ByVal hWndParent As Long, ByVal hWndChildAfter As Long, ByVal lpszClass As Long, ByVal lpszWindow As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
' TASK-P1-030: this DLL's own path, so mcp_info can hash the file that is loaded.
Private Declare Function OwnModuleFileNameW Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
Private m_cachedStart As Object
Private m_cachedBreak As Object
Private m_cachedEnd As Object
Private m_cachedFull As Object
' Card TASK-P1-005: the step commands had NO cache fallback (run/break/end did).
Private m_cachedStepInto As Object
Private m_cachedStepOver As Object
Private m_cachedStepOut As Object
Private m_cachedRtc As Object
' TASK-P1-066: state for the re-entrancy gate and the nested-call cap. These live in the
' declarations section because VB6 only accepts module-level declarations before the first
' procedure.
Private Const CALL_DEPTH_MAX As Long = 8

Private m_callDepth As Long
Private m_mutBusy As Boolean
Private m_mutBusyTool As String
Private m_mutBusyAt As Long
Private m_callToolName As String
Private m_callStartAt As Long
Private m_callIsMut As Boolean
' TASK-P1-126 (MCP-08): a COOPERATIVE cancel request. A re-entrant READ-ONLY call
' (served at the pump point while an outer call is in flight) can set this flag; the
' outer call's own pump loops call MaybeHonorCancel and abort with a structured
' {ok:false, reason:"cancelled_by_request"} receipt instead of running to
' completion. That is the only abort a single-threaded add-in can offer for a call
' already running inside the host; a call blocked inside a HOST modal loop never
' reaches a pump point of ours - use the hub tool ide_restart for that case.
Private m_cancelRequested As Boolean
Private m_cancelAt As Long
Private m_cancelByTool As String
Private m_cancelTargetTool As String
Private m_cancelTargetKind As String
Private Const CANCEL_HONOR_MIN_MS As Long = 1500
' TASK-P1-126 (MCP-09): how many read-only receipts carried the inflight block in
' this process - a cheap, machine-checkable proof that the block is really emitted.
Private m_inflightStamped As Long
' =====================================================================
' TASK-P1-168 (C / defect #10): the gate is now FORCE-RELEASABLE.
'   Measured on its own instance with the v1.8.0 build (raw receipts under
'   docs/evidence/TASK-P1-168/): a raising statement sent through
'   vba6_probe{exec_probe} does NOT make EbExecuteLine return. The IDE shows its own
'   modal run-time-error box, the single-flight gate stayed held, and every
'   write-class call was refused with reason=busy for as long as that box was up.
'   Reads kept being served (ide_state answered in 19 ms), which is the whole reason
'   this abort has to be carried by a READ-ONLY tool.
'   What this adds:
'     * m_callSeq / m_outerSeq - a per-call token. m_outerSeq is the token of the
'       OUTERMOST frame, captured at entry when the depth becomes 1, so a nested read
'       can no longer mistake ITSELF for the call that is in flight.
'     * m_abandonSeq - the token of a frame the gate was force-released for. Its
'       unwind is SWALLOWED exactly once (CallUnwind), so no frame can ever decrement
'       the depth twice.
'     * m_gateRelease* - the machine-readable record of the last forced release.
'   The release is HONEST about what it does NOT do: it does not interrupt the
'   blocked statement (a single-threaded add-in cannot preempt EbExecuteLine), so the
'   receipt says the blocked call is still running inside the IDE.
' =====================================================================
Private m_callSeq As Long
Private m_outerSeq As Long
Private m_outerTool As String
Private m_outerAt As Long
Private m_abandonSeq As Long
Private m_gateReleaseCount As Long
Private m_gateReleasedBy As String
Private m_gateReleasedTool As String
Private m_gateReleasedHeldMs As Long
Private m_gateReleasedAt As String
Private m_gateReleasedReason As String
Private Const GATE_RELEASE_MIN_MS As Long = 60000
' TASK-P1-168 (B): an ARMED release. immediate_eval / immediate_run set it before they
' hand their line to EbExecuteLine; if that call never comes back, the NEXT tool call
' sees the age and force-releases the gate by itself, so one bad statement cannot leave
' the session permanently busy.
Private m_armTool As String
Private m_armMs As Long
Private m_armAt As Long
' =====================================================================
' TASK-P1-164 (A): session-level QUIET MODE.
'   While ON the add-in must not STEAL the user's foreground: control_window's
'   action=activate skips Window.SetFocus, and every receipt (StampMcp) carries a
'   quiet{} block whose advisory says "do not type in this IDE". While OFF nothing is
'   emitted and nothing is suppressed, so the receipt stream stays byte-for-byte the
'   pre-v1.9.0 one.
' =====================================================================
Private m_quietOn As Boolean
Private m_quietSince As Long
Private m_quietTtlMs As Long
Private m_quietSuppressed As Long
Private Const QUIET_TTL_DEFAULT_MS As Long = 600000
Private Const QUIET_ADVISORY As String = "quiet mode is ON - do not type in this IDE while an agent is driving it"
' TASK-P1-168 (B): the immediate_* sink. 511 is the top of VB6's file-number range and the
' generated line closes it first, so the handle is idempotent even after a failed line.
Private Const IM_SINK_MAX_BYTES As Long = 65536
Private Const IM_FILE_NUM As Long = 511
' TASK-P1-125: ToolBuild chooses the in-IDE route itself when the project is the
' one the IDE holds, and ToolMakeProject route=cli delegates to ToolBuild - this
' override breaks the cycle and pins the in-IDE path once it has been chosen.
Private m_makeRouteOverride As String
' TASK-P1-132 (D-130-01): WHERE the code-pane caption table came from on the last
' scan ("all" | "active_fallback" | "empty"), so ide_state.wheel can explain WHY a
' pane was or was not hookable instead of only reporting the symptom.
Private m_wheelCapsSource As String
' TASK-P1-132 (D-130-01): WHICH channel resolved the last caption -> CodePane lookup
' ("caption" | "component" | "active_name" | "none"). Reported so a broken hook has an
' explained cause instead of an empty table with no reason.
Private m_wheelResolveVia As String
' TASK-P1-132 (D-130-01): the REAL outcome of the last TopLine read/write, captured
' instead of being swallowed by On Error Resume Next. status: "ok" | "pane_not_found" |
' "pane_not_writable_in_break"; err carries "VB error N (text)".
Private m_wheelTopStatus As String
Private m_wheelTopErr As String
' TASK-P1-125 (2.4): how long make_project waits for the artifact to land before it
' reports pending=true. Bounded, and abandoned early when a compiler error is seen.
Private Const MAKE_LAND_WAIT_MS As Long = 15000
' TASK-P1-068: hunk cap for the multi-hunk minimal write, and the hunk count of
' the last write (reported in the receipts next to writeCalls).
Private Const MAX_HUNKS As Long = 16
Private m_lastWriteHunks As Long
' TASK-P1-070: per-PHASE timings of the last minimal write, in milliseconds.
' These are pure additions next to the existing writeMs/writeCalls/hunks, whose
' semantics do not change. They separate OUR preparation (diff = hunk generation
' + validation; join = the rest of our preparation: line split, prefix/suffix
' scan, block build, Join) from the CodeModule calls (del = inside DeleteLines,
' ins = inside InsertLines). They are INTERVAL MEASUREMENTS OF ONE CALL - they
' do not assert a cause; that is what TASK-P1-070's matrix is for.
Private m_lastDiffMs As Long
Private m_lastJoinMs As Long
Private m_lastDelMs As Long
Private m_lastInsMs As Long
Private m_lastMaxBlock As Long
' TASK-P1-161 (R4): per-call cache of the pane->disk line map (component name -> map dict).
' VBA6 keeps inline member `Attribute X.VB_UserMemId = ...` lines on disk that the code
' pane never shows, so the disk line of a pane line is the SUM of the hidden lines
' above it and not one constant offset. Reset once per top-level call in CallTool.
Private m_paneDiskMaps As Object
' TASK-P1-085: the project-directory write manifest of the CURRENT top-level call. It is
' reset once per call in CallTool (the single tool funnel), filled by the writers that can
' put a file into the project directory on the caller's behalf - today the _undo snapshot
' store - and merged into every receipt that reports projectDirWrites.
Private m_dirWrites As Collection
' TASK-P1-121 (G-1/G-5, spec P4-108 item 6): the side-effects AUDIT BLOCK of the
' CURRENT top-level call. P4-105 measured that there is NO single receipt builder every
' tool goes through (ReceiptJson is used by 5 tools; ~70 other receipts call
' JsonStringify directly), so - exactly like m_dirWrites above - these are per-call
' accumulators that SideEffectsInto()/SealedReceipt() merge into EVERY mutating tool's
' answer at the one funnel every tool does go through (CallToolInner's text writer).
' APPEND-ONLY: no existing key or value is renamed, removed or redefined, and the
' legacy projectDirWrites key is deliberately left exactly as it was.
Private m_seCreated As Collection
Private m_seRenamed As Collection
Private m_seRemoved As Collection
Private m_seFiles As Collection
Private m_seUndo As Collection
' backupsCreated lives in Util.bas (g_seBackups) because VbFmt.bas writes .bak files too.

Public Sub Init(ByVal vbe As Object)
    Set m_vbe = vbe
End Sub

Public Function ProjectNameSafe() As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    ProjectNameSafe = p.Name
End Function

Public Function ProjectFileNameSafe() As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    ProjectFileNameSafe = p.FileName
End Function

' =====================================================================
' tool metadata
' =====================================================================

Public Function BuildToolsListResult() As Object
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim tools As Object
    Set tools = New Collection

    ' ---- contract v1.1.0 batch A (TASK-P1-017): read-only debug/GUI observers ----
    ' TASK-P1-126 (MCP-08): ide_state is the carrier of the executable abort, because
    ' it is READ-ONLY and therefore never refused by the busy gate - it can reach the
    ' pump point while a write is in flight. Its receipt always carries
    ' inflight{...} (spliced in SealedReceipt) and cancel{...}.
    Dim pIs As Object
    Set pIs = CreateObject("Scripting.Dictionary")
    Set pIs("cancel_inflight") = Prop("boolean", "TASK-P1-126 (MCP-08): ask the call currently in flight to STOP. This read-only call is served even while a WRITE is busy; it sets a cooperative cancel flag and returns cancel{...}. The in-flight call aborts at its next pump point (at least 1500 ms in) with {ok:false, reason:cancelled_by_request}. A call already blocked inside a HOST modal loop reaches none of our pump points - use the hub tool ide_restart for that case.")
    Set pIs("cancel_target") = Prop("string", "Optional: the tool you believe is in flight (informational only - the receipt reports what the add-in actually finds).")
    Set pIs("release_inflight") = Prop("boolean", "TASK-P1-168 (C / defect #10): FORCE-RELEASE the single-flight gate when the call that holds it never came back. Only acts when the in-flight call is OLDER than release_min_held_ms; the receipt gives {released, releasedTool, heldMs, abandonedSeq, statementInterrupted:false, note}. Safe to send while the gate is held, because this tool is READ-ONLY and therefore never refused by the gate.")
    Set pIs("release_min_held_ms") = Prop("integer", "The age threshold for release_inflight, in ms (default 60000). The gate is NOT touched while the in-flight call is younger than this, so a healthy long call keeps the single-flight guarantee.")
    tools.Add MakeTool("ide_state", _
        "Read the VB6 IDE mode RIGHT NOW (design/run/break) together with the raw VBA6 debug-engine flag, main window title, project name and a running flag. Never cached - a mode switch is visible on the very next call. The receipt ALWAYS carries inflight{active,tool,sinceMs,serviceable,kind} (who is in flight and whether it is a read or a write) and cancel{requested,applied,targetTool,targetKind,sinceMs,reason,note} plus release{released,releasedTool,heldMs,abandonedSeq,statementInterrupted,note}; pass cancel_inflight=true to abort the call that is in flight, or release_inflight=true to FORCE-RELEASE the gate of a call that never came back (only when it is older than release_min_held_ms, default 60000).", _
        SchemaWith(pIs))

    tools.Add MakeTool("current_position", _
        "Where the engine is currently stopped, read from the IDE itself: in break mode {available:true, project, module, proc, line, text, ebMode} with line 1-based inside the module and text the raw source line; in design/run mode {available:false, reason:not_in_break} without raising. The raw selection it read is echoed as selectedLine so the claim can be cross-checked.", _
        SchemaWith(Nothing))

    tools.Add MakeTool("list_breakpoints", _
        "List the breakpoints this tool injects into source (marker lines ' @mcp-break <id> begin / Debug.Assert False / end') as {component, line, kind, enabled}. The VBIDE object model has no breakpoint API, so IDE-native breakpoints cannot be enumerated and are reported as native:null instead of being guessed.", _
        SchemaWith(Nothing))

    ' ---- contract v1.1.0 batch C (TASK-P1-019): ui_text ----
    Dim pUi As Object
    Set pUi = CreateObject("Scripting.Dictionary")
    Set pUi("hwnd") = Prop("integer", "Window handle to read (0 or omitted = resolve the first window of this process whose class matches cls)")
    Set pUi("cls") = Prop("string", "Window class name used when hwnd is 0")
    Set pUi("depth") = Prop("integer", "How many child-window levels to include (default 3)")
    tools.Add MakeTool("ui_text", _
        "Read the accessible text of a window and its subtree. Tries UI Automation first, then MSAA (AccessibleObjectFromWindow), then WM_GETTEXT. Returns {ok:true, source:'uia|msaa|wm_gettext', text:'...'} or {ok:false, reason:'no_accessible_text', evidence:'...'} plus a channels object with each channel's raw return code/error. A negative result is a legitimate delivery: it is reported with reproducible evidence and never fabricated.", _
        SchemaWith(pUi))

    Dim pFg As Object
    Set pFg = CreateObject("Scripting.Dictionary")
    Set pFg("keyword") = Prop("string", "Substring to match (case-insensitive) against window text AND window class name; empty or omitted matches everything")
    Set pFg("cls") = Prop("string", "Exact window class name to match; empty or omitted matches every class")
    Set pFg("showing") = Prop("boolean", "true = only visible windows, false = only hidden ones; omitted = no visibility filter")
    tools.Add MakeTool("find_gui", _
        "Enumerate the windows of THIS process (top-level windows plus all their children) and filter by keyword/cls/showing. Each hit reports hwnd, cls, text, showing, enabled, rect {l,t,w,h} and path (parent title chain). Read-only.", _
        SchemaWith(pFg))

    ' ---- project info / options ----
    tools.Add MakeTool("get_project_info", _
        "Get detailed information about the open VB6 project(s): name, paths, type, IDE mode (design/run/break), options (description, help file, startup object, start project), components and references.", _
        SchemaWith(Nothing))

    Dim pOpt As Object
    Set pOpt = CreateObject("Scripting.Dictionary")
    Set pOpt("description") = Prop("string", "Project description (VBProject.Description)")
    Set pOpt("help_file") = Prop("string", "Help file path (VBProject.HelpFile)")
    Set pOpt("help_context_id") = Prop("integer", "Help context id (VBProject.HelpContextID)")
    Set pOpt("startup_object") = Prop("string", "Startup object: form name, or Sub Main / (None)")
    Set pOpt("start_project") = Prop("string", "Name of the project to make the startup project")
    Set pOpt("build_file_name") = Prop("string", "Output path for make_project (VBProject.BuildFileName)")
    Set pOpt("active_project") = Prop("string", "Name of the project to make active")
    Set pOpt("read_only_mode") = Prop("boolean", "New read-only mode")
    Set pOpt("read_only_mode_set") = Prop("boolean", "Set true to apply read_only_mode")
    tools.Add MakeTool("set_project_options", _
        "Change project/IDE options: description, help file/context, startup object, startup project, build output path, active project, read-only mode. startup_object is best-effort: on this build the VBIDE StartUpObject put is rejected for every value, so the response carries startupObjectError + startupObjectWorkaround (edit Startup in the .vbp and reopen) instead of failing the whole call.", _
        SchemaWith(pOpt))

    Dim pOpen As Object
    Set pOpen = CreateObject("Scripting.Dictionary")
    Set pOpen("path") = Prop("string", "Path of the .vbp project file to open")
    Set pOpen("exclusive") = Prop("boolean", "Open exclusively (replaces current workspace, default: false)")
    tools.Add MakeTool("open_project", _
        "Open a project file (.vbp) in the IDE (VBProjects.AddFromFile). Without exclusive it is added to the current workspace.", _
        SchemaWith(pOpen, Req("path")))

    Dim pAddProj As Object
    Set pAddProj = CreateObject("Scripting.Dictionary")
    Dim pAddProj_type As Object
    Set pAddProj_type = Prop("string", "Project type: StandardExe, ActiveXExe, ActiveXDll, ActiveXControl")
    Set pAddProj_type("enum") = EnumList("StandardExe", "ActiveXExe", "ActiveXDll", "ActiveXControl")
    Set pAddProj("type") = pAddProj_type
    Set pAddProj("template_path") = Prop("string", "Optional .vbp template path (VBProjects.AddFromTemplate)")
    tools.Add MakeTool("add_project", _
        "Create a new project in the IDE (VBProjects.Add) or from a template.", _
        SchemaWith(pAddProj))

    Dim pRemProj As Object
    Set pRemProj = CreateObject("Scripting.Dictionary")
    Set pRemProj("name") = Prop("string", "Project name (default: active project)")
    tools.Add MakeTool("remove_project", _
        "Close (remove) a project from the IDE workspace.", _
        SchemaWith(pRemProj))

    ' ---- components ----
    tools.Add MakeTool("get_component_list", _
        "List all components (modules/forms/classes) of a project with procedures, kinds, line ranges, dirty state and designer info.", _
        SchemaWith(Nothing))

    Dim pCode As Object
    Set pCode = CreateObject("Scripting.Dictionary")
    Set pCode("name") = Prop("string", "Component name (case-insensitive)")
    Set pCode("include_line_numbers") = Prop("boolean", "Prefix each line with its line number")
    Set pCode("coords") = Prop("string", "ide (default) = pure code-pane text, line numbers are IDE lines; disk = prepend the file header (Attribute/VERSION), line numbers are disk lines")
    Set pCode("proc") = Prop("string", "Return only this procedure's range: {startLine,endLine,text,sha256,coords}")
    Set pCode("symbol") = Prop("string", "Return only the range of the procedure declaring this identifier (or the single line)")
    tools.Add MakeTool("get_component_code", _
        "Read the code of a project component from IDE memory. B+C batch: coords defaults to 'ide' (BREAKING - the old behaviour was disk coordinates with the file header prepended); pass coords:'disk' for the previous text/numbering. proc: or symbol: returns just that range (startLine/endLine/text/sha256/coords); an unknown name is a structured ok:false, never an exception.", _
        SchemaWith(pCode, Req("name")))

    Dim pRep As Object
    Set pRep = CreateObject("Scripting.Dictionary")
    Set pRep("name") = Prop("string", "Component name (case-insensitive)")
    Dim pRep_operation As Object
    Set pRep_operation = Prop("string", "One of: replace, insert, append, delete, set")
    Set pRep_operation("enum") = EnumList("replace", "insert", "append", "delete", "set")
    Set pRep("operation") = pRep_operation
    Set pRep("start_line") = Prop("integer", "First line for replace/insert/delete (1-based)")
    Set pRep("end_line") = Prop("integer", "Last line for replace/delete (inclusive)")
    Set pRep("code") = Prop("string", "New code text; use CRLF or LF line breaks")
    Set pRep("expected_sha256") = Prop("string", "Optional optimistic-concurrency guard: sha256 of the current pane text")
    Set pRep("validate") = PropValidate()
    ' TASK-P1-087: opt-in pre-edit pane snapshot. Default OFF, so the hottest editor keeps
    ' exactly its current write cost; pass true when the edit must be rollback-capable.
    Set pRep("snapshot") = Prop("boolean", "true = write the PRE-EDIT pane text into the project's _undo store before this call writes anything (undo_edit can then restore it); default false = no extra file write. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("replace_code", _
        "Edit a component's code in IDE memory (line based). replace: overwrite start_line..end_line. insert: before start_line. append: after last line. delete: remove range. set: replace the whole module with code. Component is marked unsaved; use save_project. Pass snapshot=true to write a pre-edit pane snapshot that undo_edit can restore (default false = no extra file write; set already snapshots by default).", _
        SchemaWith(pRep, Req("name", "operation")))

    Dim pAddComp As Object
    Set pAddComp = CreateObject("Scripting.Dictionary")
    Dim pAddComp_type As Object
    Set pAddComp_type = Prop("string", "Component type: module, class, form, mdiform, resource, usercontrol, propPage, userdoc, designer")
    Set pAddComp_type("enum") = EnumList("module", "class", "form", "mdiform", "resource", "usercontrol", "propPage", "userdoc", "designer")
    Set pAddComp("type") = pAddComp_type
    Set pAddComp("name") = Prop("string", "Desired component name (default: IDE-generated)")
    Set pAddComp("code") = Prop("string", "Optional initial code (added via AddFromString)")
    tools.Add MakeTool("add_component", _
        "Create a new component in the active project with optional name and initial code. Standard forms fall back to a generated .frm + AddFile when VBComponents.Add(vbext_ct_MSForm) fails with E_FAIL.", _
        SchemaWith(pAddComp, Req("type")))

    Dim pRemComp As Object
    Set pRemComp = CreateObject("Scripting.Dictionary")
    Set pRemComp("name") = Prop("string", "Component name to remove")
    tools.Add MakeTool("remove_component", _
        "Remove a component from the project (VBComponents.Remove). Cannot be undone; the IDE will ask about the file.", _
        SchemaWith(pRemComp, Req("name")))

    Dim pImp As Object
    Set pImp = CreateObject("Scripting.Dictionary")
    Set pImp("path") = Prop("string", "Source file (.bas/.cls/.frm/.ctl/.dsr) or template path")
    Dim pImp_mode As Object
    Set pImp_mode = Prop("string", "file (AddFile) or template (AddFromTemplate)")
    Set pImp_mode("enum") = EnumList("file", "template")
    Set pImp("mode") = pImp_mode
    tools.Add MakeTool("import_component", _
        "Add an existing source file (or a template) to the project as a component. mode=file COPIES the source into the project directory and binds the component to that COPY (the file you pass is not linked and is verified byte-identical on return): edit the copy, not the original. componentFile and projectDirWrites name the bound copy.", _
        SchemaWith(pImp, Req("path")))

    Dim pExp As Object
    Set pExp = CreateObject("Scripting.Dictionary")
    Set pExp("name") = Prop("string", "Component name")
    Set pExp("path") = Prop("string", "Target file path (default: component's current file path)")
    tools.Add MakeTool("export_component", _
        "Save a component's source to a file. Designer components (forms/usercontrols) are exported whole via VBComponent.Export (.frm/.frx); other components write their code text.", _
        SchemaWith(pExp, Req("name")))

    Dim pRen As Object
    Set pRen = CreateObject("Scripting.Dictionary")
    Set pRen("name") = Prop("string", "Current component name")
    Set pRen("new_name") = Prop("string", "New component name")
    tools.Add MakeTool("rename_component", _
        "Rename a component in the project (VBComponent.Name).", _
        SchemaWith(pRen, Req("name", "new_name")))

    Dim pRel As Object
    Set pRel = CreateObject("Scripting.Dictionary")
    Set pRel("name") = Prop("string", "Component name to reload from disk")
    tools.Add MakeTool("reload_component", _
        "Reload a component from THE FILE IT IS BOUND TO on disk (VBComponent.Reload of FileNames(1)) - not necessarily the file you edited: import_component mode=file binds the project-local COPY, so editing the import source then reloading reports success while the code pane keeps the old text. The receipt reports reloaded (the component's real name) plus file (the bound path actually reloaded): compare file with the path you edited. Or edit in-IDE with replace_code / edit_by_anchor / apply_edits.", _
        SchemaWith(pRel, Req("name")))

    ' ---- code module ----
    Dim pAddCode As Object
    Set pAddCode = CreateObject("Scripting.Dictionary")
    Set pAddCode("component") = Prop("string", "Component name")
    Set pAddCode("code") = Prop("string", "Code text to append (CodeModule.AddFromString). Supply code or file - at least one of the two is required")
    ' TASK-P1-121 item 2 (S2/G-3), discipline 84: the description MUST stop naming
    ' AddFromFile - the file branch no longer uses it, and the observable behaviour changed
    ' (the designer head is stripped, and the target component is no longer renamed).
    Set pAddCode("file") = Prop("string", "Source file whose BODY is appended: the designer head (VERSION/BEGIN..END or Begin..End, plus the leading Attribute lines) is stripped first, so the code pane never receives a header and the target component is never renamed. Supply code or file - at least one of the two is required")
    Set pAddCode("validate") = PropValidate()
    tools.Add MakeTool("add_code", _
        "Append code to a component's code module from a string or a file. Requires code or file.", _
        SchemaWith(pAddCode, Req("component")))

    Dim pEvt As Object
    Set pEvt = CreateObject("Scripting.Dictionary")
    Set pEvt("component") = Prop("string", "Component name (usually a form/class)")
    Set pEvt("object") = Prop("string", "Name of the object raising the event (e.g. Command1)")
    Set pEvt("event") = Prop("string", "Event name (e.g. Click)")
    tools.Add MakeTool("create_event_proc", _
        "Create an event procedure stub (CodeModule.CreateEventProc), e.g. Command1_Click. Returns the line it was added at.", _
        SchemaWith(pEvt, Req("component", "object", "event")))

    Dim pFind As Object
    Set pFind = CreateObject("Scripting.Dictionary")
    Set pFind("component") = Prop("string", "Component name")
    Set pFind("target") = Prop("string", "Text to find")
    Set pFind("start_line") = Prop("integer", "Search from line (1-based, default 1)")
    Set pFind("start_col") = Prop("integer", "Search from column (1-based, default 1)")
    Set pFind("end_line") = Prop("integer", "Search up to line (default: last line)")
    Set pFind("end_col") = Prop("integer", "Search up to column (default: last column)")
    Set pFind("whole_word") = Prop("boolean", "Whole word match only")
    Set pFind("match_case") = Prop("boolean", "Case sensitive")
    Set pFind("pattern_search") = Prop("boolean", "Use pattern matching")
    tools.Add MakeTool("find_in_component", _
        "Find text inside one component using CodeModule.Find. Returns the match position (start/end line/column).", _
        SchemaWith(pFind, Req("component", "target")))

    ' ---- code pane / selection ----
    Dim pCp As Object
    Set pCp = CreateObject("Scripting.Dictionary")
    Set pCp("component") = Prop("string", "Optional component whose code pane to query (default: active code pane)")
    tools.Add MakeTool("get_codepane", _
        "Get the state of a code pane: current selection (lines/columns), top line, visible line count and view mode.", _
        SchemaWith(pCp))

    Dim pSetCp As Object
    Set pSetCp = CreateObject("Scripting.Dictionary")
    Set pSetCp("component") = Prop("string", "Optional component whose code pane to change (default: active)")
    Set pSetCp("selection") = Prop("string", "Selection as startLine,startCol,endLine,endCol (e.g. 1,1,1,1)")
    Set pSetCp("top_line") = Prop("integer", "Scroll so this line is on top")
    ' TASK-P1-114: the \`view\` property was REMOVED from this schema. VB6EXT v5.3 declares
    ' CodePane.CodePaneView GET-ONLY (the type library put-map of _CodePane holds only
    ' TopLine) and has no \`View\` property at all, and a REAL-IDE scan of all 34 command
    ' bars found NO view toggle - so the enum full_module/procedure was a promise no
    ' add-in can keep. A caller who still passes \`view\` now gets ignoredArgs.
    Set pSetCp("activate") = Prop("boolean", "Bring the code pane to the front (Show)")
    tools.Add MakeTool("set_codepane", _
        "Change a code pane: set selection/cursor, scroll and/or activate it. There is NO view-mode parameter: VB6's extensibility model exposes CodePane.CodePaneView as READ-ONLY and its command bars carry no view toggle, so an add-in cannot switch between full_module and procedure view; a \`view\` argument you pass is reported in ignoredArgs.", _
        SchemaWith(pSetCp))

    tools.Add MakeTool("get_selection", _
        "What is currently selected in the IDE: selected component, active project, active window and cursor position.", _
        SchemaWith(Nothing))

    Dim pAct As Object
    Set pAct = CreateObject("Scripting.Dictionary")
    Set pAct("name") = Prop("string", "Component name")
    Dim pAct_window As Object
    Set pAct_window = Prop("string", "code (open code pane) or designer (open designer window); default: activate component")
    Set pAct_window("enum") = EnumList("code", "designer", "activate")
    Set pAct("window") = pAct_window
    tools.Add MakeTool("activate_component", _
        "Bring a component to the front: open its code pane or designer window.", _
        SchemaWith(pAct, Req("name")))

    ' ---- files ----
    Dim pList As Object
    Set pList = CreateObject("Scripting.Dictionary")
    Set pList("path") = Prop("string", "Directory to list; relative to the project directory when not absolute (default: project directory)")
    Set pList("pattern") = Prop("string", "Semicolon-separated patterns, e.g. *.bas;*.frm (default: *)")
    Set pList("recursive") = Prop("boolean", "Recurse into subdirectories (default: true)")
    Set pList("max_entries") = Prop("integer", "Maximum number of entries (default: 500)")
    tools.Add MakeTool("list_files", _
        "List files below a directory. Returns relative paths. Skips .git/.svn/.hg/node_modules.", _
        SchemaWith(pList))

    Dim pRead As Object
    Set pRead = CreateObject("Scripting.Dictionary")
    Set pRead("path") = Prop("string", "File path; relative to the project directory when not absolute")
    Set pRead("max_lines") = Prop("integer", "Return only the first N lines (0 = all, default)")
    tools.Add MakeTool("read_file", _
        "Read a text file from disk (UTF-8 / UTF-16 / ANSI auto-detected) and return its content.", _
        SchemaWith(pRead, Req("path")))

    Dim pWrite As Object
    Set pWrite = CreateObject("Scripting.Dictionary")
    Set pWrite("path") = Prop("string", "File path; relative to the project directory when not absolute")
    Set pWrite("content") = Prop("string", "Full new content of the file")
    Dim pWrite_encoding As Object
    Set pWrite_encoding = Prop("string", "auto (ANSI for VB source files, UTF-8 otherwise), utf-8, ansi or utf-16")
    Set pWrite_encoding("enum") = EnumList("auto", "utf-8", "ansi", "utf-16")
    Set pWrite("encoding") = pWrite_encoding
    Set pWrite("create_dirs") = Prop("boolean", "Create missing directories (default: true)")
    Set pWrite("backup") = Prop("boolean", "Keep a .bak copy of the previous file (default: true)")
    Set pWrite("validate") = PropValidate()
    tools.Add MakeTool("write_file", _
        "Write a text file to disk with the given encoding. For files of the open project use reload_component afterwards (or replace_code for in-memory edits).", _
        SchemaWith(pWrite, Req("path", "content")))

    Dim pSearch As Object
    Set pSearch = CreateObject("Scripting.Dictionary")
    Set pSearch("query") = Prop("string", "Text or regular expression to search for")
    Set pSearch("path") = Prop("string", "Directory to search; relative to the project directory when not absolute (default: project directory)")
    Set pSearch("pattern") = Prop("string", "Semicolon-separated file patterns (default: VB source files)")
    Set pSearch("is_regex") = Prop("boolean", "Treat query as a regular expression (default: false)")
    Set pSearch("case_sensitive") = Prop("boolean", "Case-sensitive matching (default: false)")
    Set pSearch("max_results") = Prop("integer", "Maximum number of matches (default: 200)")
    tools.Add MakeTool("search_in_files", _
        "Search text or regex across files below a directory. Returns file, line number and line text for each match.", _
        SchemaWith(pSearch, Req("query")))

    ' ---- build / run ----
    Dim pSave As Object
    Set pSave = CreateObject("Scripting.Dictionary")
    Set pSave("path") = Prop("string", "Optional .vbp path: an unsaved project is saved there and unsaved components are saved next to it (bas/cls/frm/res by type).")
    tools.Add MakeTool("save_project", _
        "Save the project and all modified components to disk.", _
        SchemaWith(pSave))

    Dim pMake As Object
    Set pMake = CreateObject("Scripting.Dictionary")
    Set pMake("output_path") = Prop("string", "Optional output path (sets VBProject.BuildFileName before compiling)")
    Set pMake("route") = Prop("string", "ide (default) = compile inside the VB6 IDE via VBProject.MakeCompiledFile; cli = run an isolated VB6.EXE /make <vbp> /out <log> subprocess, the route that yields a REAL compiler log (the build tool's receipt shape is returned); auto = same as ide here, because make_project always acts on the project the IDE has open.")
    Set pMake("out_path") = Prop("string", "route=cli: compiler log path (default %TEMP%\vb6mcp_build.log)")
    Set pMake("timeout_ms") = Prop("integer", "route=cli: bounds ONLY the VB6.EXE /make subprocess (default 180000)")
    Set pMake("allow_host_modal") = Prop("boolean", "Default false. When VB6's own Make Project box is already up, confirming it starts a long host-modal compile that this single-threaded add-in cannot interrupt - every later write tool would then answer reason=busy until the IDE is restarted. false = leave the box alone and report modal_blocked=true with its hwnd/title/controls instead of clicking it.")
    Set pMake("laa") = Prop("boolean", "TASK-P1-132 (A10), default TRUE. After the compile, set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header. The add-in rewrites the PE header ITSELF, in place, and re-reads it to verify; an external EditBin.exe is NOT required. The receipt always reports laaBefore/laaAfter/laaSet (+ laa{} with the detail); an artifact that already carries the flag is left untouched (laaSet=false, laaAlready=true). The decision PRECEDENCE is: this per-call laa argument > the [VB6MCP] section of the project .vbp (LAA=1, optional LAANote=...) > the optional sidecar <vbpDir>\.vb6mcp.json {""laa"": true|false} > the environment variable VB6MCP_LAA > the global default (ON). The receipt reports laaSource as param | vb-project | project-file | env | global-default (plus laaProjectFile). An unreadable/invalid sidecar is NOT silently defaulted: the step is skipped and laaError/laaSource=project-file-error is returned." & _
            " WHERE the switch is stored: the named section [VB6MCP] of the .vbp is the PRIMARY source (level 2). Survival rule: a named section present WHEN THE PROJECT IS LOADED is re-emitted by VB6 on every save and survives; a section APPENDED externally WHILE the project is open is overwritten by that save - so write the .vbp with the IDE CLOSED, or use the sidecar file, which can be written at any time. It is the ROOT-area keys that the in-memory model overwrites (CondComp is that class), NOT a named section.")
    Set pMake("diagnose") = Prop("boolean", "TASK-P1-132 (MCP-lead-2), default false. Run an ISOLATED VB6.EXE /make over a COPY of the .vbp plus its component files in a temp directory and return the real compiler log plus firstError{file,line,text}. This is the only route that can diagnose a project the IDE holds (an isolated /make on the HELD project exits 0 and writes no log). It runs AUTOMATICALLY when the in-IDE Make returns compiled=true but artifactChanged=false, because that verdict is otherwise indistinguishable from a compile failure.")
    tools.Add MakeTool("make_project", _
        "Compile the project. NEVER deletes an existing artifact: the previous output is moved aside first and restored if the build fails or if the host deletes it (output{exists,size,mtime,sha256} + artifact_restored report every case, so a failed build can no longer leave you with no artifact). compiled:true alone is NOT evidence that anything was built - read compiled_means, artifactChanged and verdict. The default ide route writes NO compiler log, so log/logPath/errorCount/firstError are always present with logAvailable=false plus the reason; route=cli returns the build tool's receipt, which carries the real log, errorCount and firstError{file,line,text}." & _
        " TASK-P1-132: after a successful compile the artifact's PE header gets IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) written by this add-in itself (laa, default true; no external EditBin.exe is needed) and the receipt reports laaBefore/laaAfter/laaSet; diagnose=true compiles a COPY of the project in a temp directory with an isolated VB6.EXE /make and returns a real log plus firstError (it also runs automatically when the in-IDE Make says compiled=true but artifactChanged=false, a verdict that is otherwise indistinguishable from a compile failure).", _
        SchemaWith(pMake))

    Dim pExec As Object
    Set pExec = CreateObject("Scripting.Dictionary")
    Dim pExec_mode As Object
    Set pExec_mode = Prop("string", "run (start, or continue when already broken) | full_compile | break (pause) | continue | resume (= continue) | step_into | step_over | step_out | run_to_cursor | end | stop (= end) | restart (end then run). Aliases accepted for compatibility: stepinto, stepover, stepout, runtocursor, fullcompile.")
    Set pExec_mode("enum") = EnumList("run", "full_compile", "break", "continue", "resume", "stop", "step_into", "step_over", "step_out", "run_to_cursor", "end", "restart")
    Set pExec("mode") = pExec_mode
    tools.Add MakeTool("control_execution", _
        "Control program execution in the VB6 IDE: start/continue, start with full compile, break, single-step (into/over/out), run to cursor, end or restart. Locale independent: menu items are matched by caption, ignoring the localized bar name, across English and Chinese IDEs - including the Continue item that replaces Start while the program runs or is broken.", _
        SchemaWith(pExec, Req("mode")))

    ' ---- test host (T5 C8-C11: TestHost.bas existed but had no tool) ----
    Dim pTh As Object
    Set pTh = CreateObject("Scripting.Dictionary")
    Dim pTh_action As Object
    Set pTh_action = Prop("string", "generate (write the host .vbp) or results (score PASS/FAIL from a result file)")
    Set pTh_action("enum") = EnumList("generate", "results")
    Set pTh("action") = pTh_action
    Set pTh("vbp_path") = Prop("string", "Output .vbp of the generated test host (action=generate)")
    Set pTh("module_name") = Prop("string", "Component name in the generated vbp; defaults to the Attribute VB_Name of the module under test")
    Set pTh("module_under_test") = Prop("string", "Path of the module under test (action=generate)")
    Set pTh("assertions_path") = Prop("string", "Path of the assertions module (action=generate)")
    Set pTh("runner_path") = Prop("string", "Path of the test runner module (action=generate)")
    Set pTh("exe_name") = Prop("string", "ExeName32 for the generated host (default vb6testhost.exe)")
    Set pTh("result_path") = Prop("string", "Result file to score (action=results)")
    tools.Add MakeTool("test_host", _
        "Generate and score the offline test host (T5): writes a host .vbp around the module under test plus an assertions/runner module, or counts PASS/FAIL from the result file of the host. Closes the generate / build / run / score loop that had no tool entry point.", _
        SchemaWith(pTh, Req("action")))

    ' ---- references ----
    Dim pRef As Object
    Set pRef = CreateObject("Scripting.Dictionary")
    Set pRef("guid") = Prop("string", "Type library GUID, e.g. {00020430-0000-0000-C000-000000000046}. Required whenever mode is not file (mode defaults to guid)")
    Set pRef("path") = Prop("string", "Type library file path. Required when mode=file")
    Dim pRef_mode As Object
    Set pRef_mode = Prop("string", "guid (AddFromGuid) or file (AddFromFile)")
    Set pRef_mode("enum") = EnumList("guid", "file")
    Set pRef("mode") = pRef_mode
    Set pRef("major") = Prop("integer", "Major version (default: 0)")
    Set pRef("minor") = Prop("integer", "Minor version (default: 0)")
    tools.Add MakeTool("add_reference", _
        "Add a COM type library reference to the project by GUID or by file path.", _
        SchemaWith(pRef))

    Dim pRemRef As Object
    Set pRemRef = CreateObject("Scripting.Dictionary")
    Set pRemRef("name") = Prop("string", "Reference name (or GUID) to remove, e.g. Scripting")
    tools.Add MakeTool("remove_reference", _
        "Remove a reference from the project by name or GUID.", _
        SchemaWith(pRemRef, Req("name")))

    ' ---- design-time properties ----
    Dim pProps As Object
    Set pProps = CreateObject("Scripting.Dictionary")
    Set pProps("name") = Prop("string", "Component name (form/class/control host)")
    tools.Add MakeTool("get_component_properties", _
        "List the design-time properties of a component (the Properties window content): name, value, number of indices.", _
        SchemaWith(pProps, Req("name")))

    Dim pSetProp As Object
    Set pSetProp = CreateObject("Scripting.Dictionary")
    Set pSetProp("name") = Prop("string", "Component name")
    Set pSetProp("property") = Prop("string", "Property name (e.g. Caption)")
    Set pSetProp("value") = Prop("string", "New value; numbers/booleans are parsed, otherwise treated as string")
    Set pSetProp("indices") = Prop("string", "Comma-separated indices for indexed properties (up to 4)")
    tools.Add MakeTool("set_component_property", _
        "Change a design-time property of a component (Property.Value / IndexedValue).", _
        SchemaWith(pSetProp, Req("name", "property", "value")))

    ' ---- members ----
    Dim pMem As Object
    Set pMem = CreateObject("Scripting.Dictionary")
    Set pMem("component") = Prop("string", "Component name (class module)")
    tools.Add MakeTool("list_members", _
        "List the public members of a class component (CodeModule.Members) with type, scope, static flag and code location.", _
        SchemaWith(pMem, Req("component")))

    Dim pSetMem As Object
    Set pSetMem = CreateObject("Scripting.Dictionary")
    Set pSetMem("component") = Prop("string", "Component name")
    Set pSetMem("member") = Prop("string", "Member name")
    Set pSetMem("attribute") = Prop("string", "Attribute: description, help_context_id, hidden, browsable, ui_default, bindable, default_bind, display_bind, request_edit, category, property_page, standard_method")
    Set pSetMem("value") = Prop("string", "New value (booleans parsed from true/false/1/0)")
    tools.Add MakeTool("set_member_attribute", _
        "Change an attrName of a class member: description, help context, hidden, browsable, UIDefault, data binding flags, category, property page, standard method.", _
        SchemaWith(pSetMem, Req("component", "member", "attribute", "value")))

    ' ---- windows ----
    tools.Add MakeTool("list_windows", _
        "List all IDE windows with caption, kind, visibility, state, position and size.", _
        SchemaWith(Nothing))

    Dim pWin As Object
    Set pWin = CreateObject("Scripting.Dictionary")
    Set pWin("window") = Prop("string", "Window caption (substring) or 1-based index")
    Dim pWin_action As Object
    Set pWin_action = Prop("string", "activate, close, show, hide, minimize, maximize, restore, move, resize, set_caption")
    Set pWin_action("enum") = EnumList("activate", "close", "show", "hide", "minimize", "maximize", "restore", "move", "resize", "set_caption")
    Set pWin("action") = pWin_action
    Set pWin("caption") = Prop("string", "New caption (read-only in VB6EXT; action accepted but ignored; whenever it is passed, the receipt reports it in ignoredArgs)")
    Set pWin("left") = Prop("integer", "New left coordinate (action=move)")
    Set pWin("top") = Prop("integer", "New top coordinate (action=move)")
    Set pWin("width") = Prop("integer", "New width (action=resize)")
    Set pWin("height") = Prop("integer", "New height (action=resize)")
    tools.Add MakeTool("control_window", _
        "Control an IDE window: focus, close, show/hide, minimize/maximize/restore, move, resize, rename.", _
        SchemaWith(pWin, Req("window", "action")))

    ' ---- form designer ----
    Dim pForm As Object
    Set pForm = CreateObject("Scripting.Dictionary")
    Set pForm("form") = Prop("string", "Form component name")
    Dim pForm_action As Object
    Set pForm_action = Prop("string", "info, select_all, clear_selection, cut, copy, paste")
    Set pForm_action("enum") = EnumList("info", "props", "menu_nest", "select_all", "clear_selection", "cut", "copy", "paste")
    Set pForm("action") = pForm_action
    Set pForm("control") = Prop("string", "Control NAME (action=props or menu_nest; preferred over index)")
    Set pForm("index") = Prop("integer", "1-based control index (action=props; see form_designer info)")
    Set pForm("parent") = Prop("string", "action=menu_nest: parent VB.Menu name (empty = make it a top-level item again)")
    tools.Add MakeTool("form_designer", _
        "Work with the form designer: info lists controls (name/class/progid/index/selection), props dumps one control's design-time property values (by name or 1-based index), menu_nest nests a VB.Menu item inside another menu item (VB6 has no menu-level property - the tree lives in the .frm indentation, so the form is flushed to disk, the block is re-nested, and the component is reloaded), or applies selection/clipboard operations.", _
        SchemaWith(pForm, Req("form", "action")))

    Dim pAddCtl As Object
    Set pAddCtl = CreateObject("Scripting.Dictionary")
    Set pAddCtl("form") = Prop("string", "Form component name")
    Set pAddCtl("prog_id") = Prop("string", "Control ProgID, e.g. VB.CommandButton")
    Set pAddCtl("name") = Prop("string", "Desired control name (default: the IDE-generated name)")
    tools.Add MakeTool("add_control", _
        "Add a control to a form at design time (VBControls.Add by ProgID). The response echoes the resolved control name to use with set_control_property/remove_control. The part is PRE-VALIDATED against the project@s registered control parts: VBControls.Add hangs the whole IDE with no dialog and no timeout when the class is not a component of this project, so an unregistered ProgID is refused immediately instead.", _
        SchemaWith(pAddCtl, Req("form", "prog_id")))

    tools.Add MakeTool("list_component_parts", _
        "Read-only: list the ActiveX control parts this project has registered (its .vbp Object= lines) as {guid, version, path, resolvedPath, exists}. Intrinsic VB.* controls are compiled in, need no declaration and are therefore NOT listed; unknown ProgIDs are not resolved to a part, so no progIds are reported.", _
        SchemaWith(Nothing))

    Dim pRemCtl As Object
    Set pRemCtl = CreateObject("Scripting.Dictionary")
    Set pRemCtl("form") = Prop("string", "Form component name")
    Set pRemCtl("index") = Prop("integer", "1-based index of the control (see form_designer info)")
    Set pRemCtl("control") = Prop("string", "Control NAME (preferred over index; see form_designer info)")
    tools.Add MakeTool("remove_control", _
        "Remove a control from a form at design time (VBControls.Remove). Address it by control name (preferred) or by 1-based index.", _
        SchemaWith(pRemCtl, Req("form")))

    Dim pSetCtl As Object
    Set pSetCtl = CreateObject("Scripting.Dictionary")
    Set pSetCtl("form") = Prop("string", "Form component name")
    Set pSetCtl("index") = Prop("integer", "1-based index of the control")
    Set pSetCtl("control") = Prop("string", "Control NAME (preferred over index; see form_designer info)")
    Set pSetCtl("property") = Prop("string", "Design-time property name (e.g. Caption)")
    Set pSetCtl("value") = Prop("string", "New value; numbers/booleans parsed, otherwise string")
    tools.Add MakeTool("set_control_property", _
        "Change a design-time property of a control on a form. Address it by control name (preferred) or by 1-based index; the response echoes the resolved control and index.", _
        SchemaWith(pSetCtl, Req("form", "property", "value")))

    ' ---- property bag ----
    Dim pBag As Object
    Set pBag = CreateObject("Scripting.Dictionary")
    Dim pBag_target As Object
    Set pBag_target = Prop("string", "project or component")
    Set pBag_target("enum") = EnumList("project", "component")
    Set pBag("target") = pBag_target
    Set pBag("name") = Prop("string", "Component name (target=component)")
    Set pBag("section") = Prop("string", "Section (project property bag only; with target=component it is reported in ignoredArgs)")
    Set pBag("key") = Prop("string", "Key to read or write")
    Set pBag("value") = Prop("string", "Value to write (omit to read): its PRESENCE selects a WRITE; omit it to read")
    Set pBag("has_value") = Prop("boolean", "Explicit write trigger (default false): true writes value even when it is absent or empty. When omitted it is inferred from the presence of value")
    tools.Add MakeTool("property_bag", _
        "Read/write named values in a project's or component's property bag (ReadProperty/WriteProperty).", _
        SchemaWith(pBag, Req("target", "key")))

    ' ---- add-ins / command bars / quit ----
    tools.Add MakeTool("list_addins", _
        "List all registered add-ins with their connection state.", _
        SchemaWith(Nothing))

    Dim pAddin As Object
    Set pAddin = CreateObject("Scripting.Dictionary")
    Set pAddin("prog_id") = Prop("string", "Add-in ProgID (e.g. VB6MCP.Connect)")
    Set pAddin("connect") = Prop("boolean", "True to load/enable, false to unload/disable")
    tools.Add MakeTool("set_addin_state", _
        "Load/unload an add-in (AddIn.Connect).", _
        SchemaWith(pAddin, Req("prog_id", "connect")))

    Dim pCb As Object
    Set pCb = CreateObject("Scripting.Dictionary")
    Dim pCb_action As Object
    Set pCb_action = Prop("string", "list (all bars and their controls, with id/enabled/visible/path) or execute (run exactly one menu command)")
    Set pCb_action("enum") = EnumList("list", "execute")
    Set pCb("action") = pCb_action
    Set pCb("bar") = Prop("string", "Optional bar filter: a bar name OR a menu path segment (matched as a caption prefix), e.g. ""调试"" or ""Debug""")
    Set pCb("control") = Prop("string", "Required for action=execute. A menu command caption (prefix search allowed), or an id as ""#746"" / ""id:746"" / a bare number, optionally ""#746 <caption>"" to disambiguate.")
    tools.Add MakeTool("command_bars", _
        "List the IDE menus (CommandBars) or execute a menu command. action=list returns every bar and control with its language-independent identity (id, tag, type, shortcutText, enabled, visible, path). action=execute runs one command, addressed either by caption or by id as ""#746"" / ""id:746"" / a bare number; because VB6 ids are NOT unique, an id-only lookup executes only when EXACTLY ONE ENABLED control matches and otherwise returns a structured id_ambiguous failure with idCandidates - disambiguate with bar= or by appending the caption (""#746 <caption>""). The CommandBars model is suspended while a program runs, so use this in design or break state.", _
        SchemaWith(pCb, Req("action")))

    ' ---- contract v1.5.0 (TASK-P1-069): message-only dialog primitives ----
    ' The three tools below are the "one call to escape a modal box" primitives.
    ' They ONLY send window messages / enumerate windows: no WriteProcessMemory,
    ' no injection, no IAT patch. read_active_dialogs is read-only;
    ' close_active_dialogs is EXPLICITLY INVOKED ONLY and clicks a whitelisted
    ' button caption; send_window_message refuses hwnds of other processes.
    Dim pDlg As Object
    Set pDlg = CreateObject("Scripting.Dictionary")
    Set pDlg("cls") = Prop("string", "Dialog window class to match (default #32770 = a Windows dialog); ""*"" = every visible top-level window of this process")
    Set pDlg("include_children") = Prop("boolean", "Include the child controls and the clickable button list (default true)")
    Set pDlg("max") = Prop("integer", "Maximum number of dialogs to return (default: all)")
    tools.Add MakeTool("read_active_dialogs", _
        "Read-only: enumerate the ACTIVE dialogs of THIS VB6 process (default window class #32770) as {hwnd, cls, title, rect, owner, ownerTitle, ownerDisabled, foreground, childCount, buttons[], controls[]}. ownerDisabled=true is the reliable signal that the box is BLOCKING the IDE - a merely visible window is not enough. Enumerates and reads text only; never sends a message.", _
        SchemaWith(pDlg))

    Dim pDlgClose As Object
    Set pDlgClose = CreateObject("Scripting.Dictionary")
    Set pDlgClose("hwnd") = Prop("integer", "Target dialog hwnd (0/omitted = every active dialog)")
    Set pDlgClose("button") = Prop("string", "Only click this button caption (it must still pass the safe whitelist)")
    Set pDlgClose("dry_run") = Prop("boolean", "true = report what WOULD be clicked without sending anything; a dry run returns exactly the same key set as a real click, host_modal_risk included")
    Set pDlgClose("allow_host_modal") = Prop("boolean", "Default false. A dialog whose own button starts a long VB6-modal operation (VB6's Make Project box, detected from its title or from its Save &in: / Save as &type: / File &name: controls) is reported as host_modal_risk=known and is NOT clicked, because the click leaves the host inside a modal loop that this single-threaded add-in cannot interrupt - every later write tool would then answer reason=busy (busyTool=close_active_dialogs) until the IDE is restarted. Set true to click it anyway.")
    tools.Add MakeTool("close_active_dialogs", _
        "EXPLICITLY INVOKED ONLY - nothing in this server calls it automatically. Click a SAFE button of an active dialog (default #32770) with the BM_CLICK window message: the caption must be on the safe whitelist (OK/Yes/Cancel/Close/Continue/Ignore + Chinese twins) and never on the danger denylist (delete/remove/uninstall/format/overwrite/discard/reset/...). Disabled or non-whitelisted buttons are reported as skipped and are never clicked. " _
        & "SAFETY: before clicking, the dialog BODY (title + every child control's text) is READ; if it names a destructive action the DEFAULT click is REFUSED with reason=dangerous_dialog_body + bodyHit (an explicit button=""No"" is the escape hatch) and NOTHING is sent. " _
        & "TASK-P1-125 adds a second, independent guard: every path - the dry_run preview included - carries host_modal_risk (none|possible|known) + host_modal_risk_reason + host_modal_risk_class + riskDialogs[], and a box whose button would start a host-modal operation is NOT clicked unless allow_host_modal=true; a Cancel/Close caption is still allowed on such a box, because cancelling cannot start the operation. recovery_hint explains what to do if a write tool then answers reason=busy. Every path returns the same keys: clickedCount, dialogsBefore, dialogsAfter, clicked[].stillExists, skipped[], remaining[].", _
        SchemaWith(pDlgClose))

    Dim pSend As Object
    Set pSend = CreateObject("Scripting.Dictionary")
    Set pSend("hwnd") = Prop("integer", "Target window handle - must belong to THIS VB6 process")
    Set pSend("action") = Prop("string", "click | key | message (default message)")
    Set pSend("message") = Prop("string", "action=message: a name (click bm_click wm_command wm_keydown wm_keyup wm_char wm_close wm_mousewheel wm_vscroll wm_syscommand bm_setstate cb_showdropdown em_linescroll) or a number such as 0x0111")
    Set pSend("vk") = Prop("string", "action=key: a VK name (ENTER ESC TAB F1..F12 UP DOWN LEFT RIGHT A..Z 0..9) or a number such as 0x0D")
    Set pSend("wparam") = Prop("string", "wParam (decimal or 0x hex)")
    Set pSend("lparam") = Prop("string", "lParam (decimal or 0x hex)")
    Set pSend("presses") = Prop("integer", "How many times to send it (default 1, max 20)")
    Set pSend("post") = Prop("boolean", "true = PostMessage (async) instead of SendMessage")
    Set pSend("confirm") = Prop("boolean", "action=message ONLY: HIGH-RISK messages (WM_CLOSE 0x0010 / WM_QUIT 0x0012 / WM_SYSCOMMAND 0x0112) are REFUSED with reason=high_risk_message unless confirm=true (default false). Nothing is sent when refused.")
    tools.Add MakeTool("send_window_message", _
        "Send a window message to a hwnd of THIS VB6 process. The parameter is named message (NOT msg: an argument called msg is ignored and the call fails with reason=bad_message - a real caller trap). action=click (BM_CLICK); action=key (WM_KEYDOWN+WM_KEYUP for the vk name/number); action=message takes a NAME (click bm_click wm_command wm_keydown wm_keyup wm_char wm_close wm_mousewheel wm_vscroll wm_syscommand bm_setstate cb_showdropdown em_linescroll) OR A RAW NUMBER such as 0x0111 / 522 / &H20A. SendMessage by default, post=true for PostMessage. HIGH-RISK messages (WM_CLOSE 0x0010 / WM_QUIT 0x0012 / WM_SYSCOMMAND 0x0112) are refused unless confirm=true. MESSAGE ONLY: no memory is written, nothing is injected and no IAT is patched; a hwnd of another process (and HWND_BROADCAST) is refused.", _
        SchemaWith(pSend, Req("hwnd")))

    tools.Add MakeTool("quit_ide", _
        "Quit the VB6 IDE (VBE.Quit). Unsaved changes are lost - save first!", _
        SchemaWith(Nothing))

    Dim pVal As Object
    Set pVal = CreateObject("Scripting.Dictionary")
    Set pVal("name") = Prop("string", "Component name (kind inference + reporting)")
    Set pVal("code") = Prop("string", "Code text to validate (NOT written to the IDE)")
    Set pVal("kind") = Prop("string", "bas|cls|frm (optional; inferred from name)")
    Set pVal("mode") = Prop("string", "off|warn|strict (default warn)")
    Set pVal("baseline") = Prop("string", "TASK-P1-126 (MCP-15): a previous run to diff against - either the stableKey lines or the WHOLE previous receipt text (both are scanned for stableKey values). The receipt then reports baseline_diff{added[],removed[]} matched on stableKey = ruleId|message|proc, so inserting lines and shifting every line number does NOT make a pre-existing diagnostic look new.")
    tools.Add MakeTool("validate_code", _
        "Validate VB6 code (lexical + structural + agent-habit rules) before writing; returns diagnostics with stable keys, line numbers and sha256. Every diagnostic carries proc + procOffset + fingerprint + stableKey (TASK-P1-126/MCP-15), and baseline_diff{added[],removed[],unchangedCount} is always present (enabled=false unless baseline is given).", _
        SchemaWith(pVal, Req("code")))

    Dim pRs As Object
    Set pRs = CreateObject("Scripting.Dictionary")
    Set pRs("path") = Prop("string", "Source file path (relative paths resolve against the project dir)")
    tools.Add MakeTool("read_source", _
        "Read a VB6 source file with encoding detection (utf-8/ansi/utf-16le), byte-level sha256 and decoded text.", _
        SchemaWith(pRs, Req("path")))

    Dim pCc As Object
    Set pCc = CreateObject("Scripting.Dictionary")
    Set pCc("name") = Prop("string", "Component name")
    tools.Add MakeTool("check_component", _
        "Lint the current in-IDE text of a component (same diagnostics as validate_code, evaluated on IDE memory).", _
        SchemaWith(pCc, Req("name")))

    Dim pB As Object
    Set pB = CreateObject("Scripting.Dictionary")
    Set pB("project") = Prop("string", "Project .vbp path (optional; default = active project)")
    Set pB("defines") = Prop("string", "Comma-separated /d defines, e.g. TESTING=1")
    Set pB("out_path") = Prop("string", "Build log path (optional)")
    Set pB("timeout_ms") = Prop("integer", "Timeout in ms (default 180000)")
    Set pB("laa") = Prop("boolean", "TASK-P1-132 (A10), default TRUE. After the compile, set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header. The add-in rewrites the PE header ITSELF, in place, and re-reads it to verify; an external EditBin.exe is NOT required. The receipt always reports laaBefore/laaAfter/laaSet (+ laa{} with the detail); an artifact that already carries the flag is left untouched (laaSet=false, laaAlready=true). The decision PRECEDENCE is: this per-call laa argument > the [VB6MCP] section of the project .vbp > the optional sidecar <vbpDir>\.vb6mcp.json > the environment variable VB6MCP_LAA > the global default (ON); the receipt reports laaSource (param | project-file | global-default | env) and laaProjectFile, and an invalid sidecar is reported as project-file-error instead of being silently defaulted.")
    tools.Add MakeTool("build", _
        "Compile the project with an isolated VB6.EXE /make <vbp> /out <log> subprocess (no modal) - the reliable command-line route. Returns log (the compiler log TEXT), logPath proven readable via logAvailable/logBytes, errorCount, firstError{file,line,lineSpace,text}, output{exists,size,mtime,sha256}, and artifact_locked + artifact_holder_pid when the artifact was not rewritten because a running process holds it (instead of a bare permission-denied). Note: only the FIRST error is reported per build. TASK-P1-132: laa (default true) makes the add-in set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header itself (no external EditBin.exe) and the receipt reports laaBefore/laaAfter/laaSet.", _
        SchemaWith(pB))

    Dim pEba As Object
    Set pEba = CreateObject("Scripting.Dictionary")
    Set pEba("name") = Prop("string", "Component name")
    Set pEba("anchor") = Prop("string", "Exact line to replace (must match uniquely)")
    Set pEba("new_text") = Prop("string", "Replacement line")
    Set pEba("expected_sha256") = Prop("string", "Optional optimistic-concurrency guard")
    Set pEba("occurrence") = Prop("integer", "1-based index of the hit to use when the anchor repeats; omitted = the anchor must be unique (A-batch rule)")
    Set pEba("match") = Prop("string", "normalized (default: ignore indentation and case) or exact (raw line)")
    Set pEba("scope") = Prop("object", "Optional {proc?, start_line?, end_line?} limiting where the anchor may match")
    Set pEba("validate") = PropValidate()
    ' TASK-P1-087: same opt-in knob as replace_code (this tool already snapshots through the
    ' write funnel by default; snapshot=true keeps that guarantee explicit and writes no
    ' second file).
    Set pEba("snapshot") = Prop("boolean", "true = guarantee a PRE-EDIT pane snapshot exists in the project's _undo store (this tool already writes one by default through the write funnel); default false = unchanged. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("edit_by_anchor", _
        "Replace one uniquely-matched line by content anchor (no line numbers). Rejects if the anchor is missing or ambiguous. snapshot=true is accepted (a pre-edit pane snapshot for undo_edit is written by default here).", _
        SchemaWith(pEba, Req("name", "anchor", "new_text")))

    Dim pAe As Object
    Set pAe = CreateObject("Scripting.Dictionary")
    Set pAe("name") = Prop("string", "Component name")
    Set pAe("edits") = Prop("array", "Array of {op:R|I|D, anchor, text}")
    Set pAe("expected_sha256") = Prop("string", "Optional guard; whole batch is all-or-nothing")
    Set pAe("match") = Prop("string", "normalized (default) or exact")
    Set pAe("validate") = PropValidate()
    ' TASK-P1-087: same opt-in knob as replace_code (this tool already snapshots through the
    ' write funnel by default; snapshot=true keeps that guarantee explicit and writes no
    ' second file).
    Set pAe("snapshot") = Prop("boolean", "true = guarantee a PRE-EDIT pane snapshot exists in the project's _undo store (this tool already writes one by default through the write funnel); default false = unchanged. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("apply_edits", _
        "Apply a batch of anchor edits atomically (any failure leaves the component unchanged). snapshot=true is accepted (a pre-edit pane snapshot for undo_edit is written by default here).", _
        SchemaWith(pAe, Req("name", "edits")))

    Dim pPe As Object
    Set pPe = CreateObject("Scripting.Dictionary")
    Set pPe("name") = Prop("string", "Component name")
    Set pPe("edits") = Prop("array", "Array of {op:R|I|D, anchor, text}")
    tools.Add MakeTool("preview_edit", _
        "Return a unified-style diff for a batch of anchor edits without writing anything. B+C batch (contract invariant 8): the algorithm is a LINE-LEVEL LCS with hunk grouping, so an insertion in the middle renders as ONE hunk of +N (not a delete block plus an add block); contextLines (default 3) of context are emitted around each hunk. hunks[] carries {kind,start,count,oldStart,oldCount,newStart,newCount,contextLines}; 'algorithm' reports lcs | fallback-too-large | none, and the fallback is stated instead of silently degrading. Still a preview: nothing is written.", _
        SchemaWith(pPe, Req("name", "edits")))

    Dim pWs As Object
    Set pWs = CreateObject("Scripting.Dictionary")
    Set pWs("path") = Prop("string", "Target file path")
    Set pWs("text") = Prop("string", "Text to write (ASCII string literals only for VB files)")
    Set pWs("encoding") = Prop("string", "auto|ansi|utf-8|utf-16le (default auto: VB files -> ansi)")
    Set pWs("validate") = PropValidate()
    tools.Add MakeTool("write_source", _
        "Write a VB source/text file with explicit encoding and byte-level sha256 (does not touch the IDE).", _
        SchemaWith(pWs, Req("path", "text")))

    Dim pVf As Object
    Set pVf = CreateObject("Scripting.Dictionary")
    Set pVf("name") = Prop("string", "Component name")
    Set pVf("disk_path") = Prop("string", "Optional disk file (default = component file)")
    Set pVf("baseline") = Prop("string", "Optional lint baseline keys (newline separated)")
    tools.Add MakeTool("verify", _
        "Aggregate verdict for one component: IDE==DISK + lint diagnostics + optional baseline filter.", _
        SchemaWith(pVf, Req("name")))

    Dim pIt As Object
    Set pIt = CreateObject("Scripting.Dictionary")
    Set pIt("name") = Prop("string", "Component name")
    Set pIt("procedure") = Prop("string", "Procedure to instrument")
    Set pIt("expressions") = Prop("string", "Comma-separated expressions, e.g. x,y")
    Set pIt("id") = Prop("string", "Trace id (also the removal handle)")
    Set pIt("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("inject_trace", _
        "Insert a McpTrace call at the start of a procedure. B+C/TASK-P1-028: the guard is a RUNTIME 'If McpTraceOn() Then' (not '#If MCP_TRACE'), because a compile-time condition only applies when the project is loaded with that CondComp - so this needs no CondComp and no reopen. TASK-P1-129: inject_trace now GENERATES AND ATTACHES the module modMcpTrace itself when the project does not have it (nothing to import by hand) and verifies the attachment before answering; the generated header comment language is chosen UI-FIRST (the IDE's own UI language when it can be read, else the VB6.EXE version resource, else English) and the receipt reports all three signals. The sink is enabled by the MCP_TRACE_ON / MCP_TRACE_FILE environment variables of the VB6 process (see the response hints).", _
        SchemaWith(pIt, Req("name", "procedure", "id")))

    Dim pRt As Object
    Set pRt = CreateObject("Scripting.Dictionary")
    Set pRt("name") = Prop("string", "Component name")
    Set pRt("id") = Prop("string", "Trace id to remove")
    Set pRt("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("remove_trace", _
        "Remove a previously injected trace block; sha256 returns to the pre-injection value. TASK-P1-149: the modMcpTrace component is NEVER removed automatically - the receipt always reports moduleRemoved=false, moduleLeftInPlace=true and moduleRemoveReason=""kept_by_design"" (moduleRemoveReason is module_absent when no such component exists), because (a) you may have added your own code to it and (b) other @mcp-trace blocks elsewhere may still call McpTrace/McpTraceOn, so the project would stop compiling without it. moduleRemoveNote carries the manual removal steps (verify no '@mcp-trace' block remains, then call remove_component yourself). The old automatic removal (removed_autogen_module) no longer exists.", _
        SchemaWith(pRt, Req("name", "id")))

    Dim pFi As Object
    Set pFi = CreateObject("Scripting.Dictionary")
    Set pFi("name") = Prop("string", "Component name")
    Set pFi("procedure") = Prop("string", "Procedure to fault")
    Set pFi("mode") = Prop("string", "raise|return|delay")
    Set pFi("value") = Prop("string", "Mode payload (return value / delay ms; optional)")
    Set pFi("id") = Prop("string", "Fault id (removal handle)")
    Set pFi("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("fault_inject", _
        "Inject a guarded fault (raise/return/delay) at the start of a procedure (#If MCP_FAULT).", _
        SchemaWith(pFi, Req("name", "procedure", "mode", "id")))

    Dim pCf As Object
    Set pCf = CreateObject("Scripting.Dictionary")
    Set pCf("name") = Prop("string", "Component name")
    Set pCf("id") = Prop("string", "Fault id to clear")
    Set pCf("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("clear_faults", _
        "Remove a previously injected fault block.", _
        SchemaWith(pCf, Req("name", "id")))

    Dim pAt As Object
    Set pAt = CreateObject("Scripting.Dictionary")
    Set pAt("count") = Prop("integer", "How many recent entries (default 50)")
    tools.Add MakeTool("audit_tail", _
        "Return the most recent tool-call audit entries (newest first).", _
        SchemaWith(pAt))

    Dim pSbp As Object
    Set pSbp = CreateObject("Scripting.Dictionary")
    Set pSbp("name") = Prop("string", "Component name")
    Set pSbp("anchor") = Prop("string", "Line to break before (must match uniquely)")
    Set pSbp("id") = Prop("string", "Breakpoint id (removal handle)")
    Set pSbp("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("set_breakpoint", _
        "Inject a persistent Debug.Assert False breakpoint before a uniquely-matched line (auto-omitted in compiled builds). The anchor must be a statement INSIDE a procedure body: a module-level line or a procedure declaration is refused with ok:false, reason:'anchor_outside_procedure' (injecting there puts executable code outside every procedure and the project stops compiling).", _
        SchemaWith(pSbp, Req("name", "anchor", "id")))

    Dim pCbp As Object
    Set pCbp = CreateObject("Scripting.Dictionary")
    Set pCbp("name") = Prop("string", "Component name")
    Set pCbp("id") = Prop("string", "Breakpoint id")
    Set pCbp("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("clear_breakpoints", _
        "Remove an injected breakpoint block by id.", _
        SchemaWith(pCbp, Req("name", "id")))

    Dim pDc As Object
    Set pDc = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("debug_commands", _
        "Catalog of the VB6 debug commands this server knows, for use with the CommandBars path. Per entry: id (this server's stable command NAME, e.g. step_into - NOT a CommandBar id), keys (shortcut), english (English caption), menuId (the live CommandBar Control.Id, or the measured table value when the control cannot be resolved live), shortcutText, caption (LOCALIZED - the Chinese IDE returns Chinese) and idSource (live|table). It is a FIXED list, not a live probe of the IDE: it carries no menu path and no enabled/visible state. For ids, captions, paths and enabled flags call command_bars action=list.", _
        SchemaWith(pDc))

    Dim pEs As Object
    Set pEs = CreateObject("Scripting.Dictionary")
    Set pEs("pid") = Prop("integer", "Target process id (default: the IDE process hosting this add-in)")
    tools.Add MakeTool("execution_state", _
        "Win32 execution state of a process: main window, responsive / hung, top-level window count, modal dialog count.", _
        SchemaWith(pEs))

    Dim pVp As Object
    Set pVp = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("vba6_probe_plan", _
        "Ordered plan for the VBA6.DLL debug-engine probes (read-only first, writes last) with preconditions. Run only inside the IDE process.", _
        SchemaWith(pVp))

    Dim pEa As Object
    Set pEa = CreateObject("Scripting.Dictionary")
    Set pEa("name") = Prop("string", "Component name")
    tools.Add MakeTool("export_public_api", _
        "Freeze the component public API as a spec (JSON members[] plus a TSV spec field). Text level; parameter names are dropped.", _
        SchemaWith(pEa))

    Dim pCa As Object
    Set pCa = CreateObject("Scripting.Dictionary")
    Set pCa("name") = Prop("string", "Component name")
    Set pCa("baseline") = Prop("string", "Baseline spec TSV previously returned by export_public_api (its spec field)")
    ' TASK-P1-109 (D-1107-02): the implementation at ToolCheckPublicApi REQUIRES a non-empty
    ' baseline (Err.Raise "baseline spec is required"), and TASK-P1-108's central classifier already
    ' answers that condition as reason:"required_argument_missing" / argument:"baseline". The schema,
    ' however, declared NO "required" key at all, so a schema-validating caller could omit baseline
    ' and only learn about it from a runtime refusal. DIRECTION CHOSEN: DECLARE it (below).
    ' Rejected alternative - "accept an empty baseline": it would have to INVENT a success (every
    ' current member reported as added[]), i.e. exactly the misleading-success class this product
    ' treats as a defect (D-1107-04/05), and it would delete the classification P1-108 delivered.
    ' NO BEHAVIOUR CHANGE: a missing baseline was refused before and is refused now; only the
    ' DECLARATION moves (see 内部证据 section (1)).
    tools.Add MakeTool("check_public_api", _
        "Compare the current component public API against a frozen spec; returns violations[] (removed / changed) and added[].", _
        SchemaWith(pCa, Req("baseline")))

    Dim pDg As Object
    Set pDg = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("dependency_graph", _
        "Text level dependency graph of the active project: nodes, edges (component name references), cycles, and a purity classification per component.", _
        SchemaWith(pDg))

    Dim pPu As Object
    Set pPu = CreateObject("Scripting.Dictionary")
    Set pPu("name") = Prop("string", "Component name")
    tools.Add MakeTool("purity_check", _
        "Purity classification of one component (can it go into an offline test host): reasons such as new / createobject / fileio / apideclare / shell / gui.", _
        SchemaWith(pPu))

    Dim pGs As Object
    Set pGs = CreateObject("Scripting.Dictionary")
    Set pGs("name") = Prop("string", "Component name")
    tools.Add MakeTool("generate_stub", _
        "Generate a compilable stub module from a component: public declarations kept verbatim, public bodies emptied, private members dropped.", _
        SchemaWith(pGs))

    Dim pRf As Object
    Set pRf = CreateObject("Scripting.Dictionary")
    Set pRf("path") = Prop("string", "Fixture file path")
    Set pRf("text") = Prop("string", "Sample text to record")
    tools.Add MakeTool("record_fixture", _
        "Record a deterministic sample (fixture) as raw ANSI bytes for later regression comparison.", _
        SchemaWith(pRf, Req("path")))

    Dim pKf As Object
    Set pKf = CreateObject("Scripting.Dictionary")
    Set pKf("path") = Prop("string", "Fixture file path")
    Set pKf("text") = Prop("string", "Current text to compare against the recorded fixture")
    tools.Add MakeTool("compare_fixture", _
        "Compare text against a recorded fixture: same / added / removed / firstDiffLine.", _
        SchemaWith(pKf, Req("path")))

    ' ---- lessons library (经验库) ----
    Dim pLl As Object
    Set pLl = CreateObject("Scripting.Dictionary")
    Set pLl("includeDisabled") = Prop("boolean", "Include disabled lessons (default false)")
    tools.Add MakeTool("list_lessons", _
        "List the VB6MCP lessons (experience library): scenario / trap / do with scope and enabled state.", _
        SchemaWith(pLl))

    Dim pAl As Object
    Set pAl = CreateObject("Scripting.Dictionary")
    Set pAl("id") = Prop("string", "Stable lesson id (optional; auto custom-N)")
    Set pAl("scenario") = Prop("string", "When this applies")
    Set pAl("trap") = Prop("string", "The pitfall")
    Set pAl("do") = Prop("string", "What to do instead")
    Set pAl("scope") = Prop("string", "all | mcp | vb6 (optional)")
    Set pAl("enabled") = Prop("boolean", "Enabled (default true)")
    tools.Add MakeTool("add_lesson", _
        "Add or overwrite a lesson in the user layer (mcp_lessons/user.json) with a .bak backup.", _
        SchemaWith(pAl))

    Dim pDl As Object
    Set pDl = CreateObject("Scripting.Dictionary")
    Set pDl("id") = Prop("string", "Lesson id")
    tools.Add MakeTool("delete_lesson", _
        "Remove a user lesson, or disable a shipped one (writes enabled=false).", _
        SchemaWith(pDl))

    Dim pEl As Object
    Set pEl = CreateObject("Scripting.Dictionary")
    Set pEl("includeDisabled") = Prop("boolean", "Include disabled lessons (default false)")
    tools.Add MakeTool("export_lessons", _
        "Export the merged lesson library as Markdown (re-importable via import_lessons_md).", _
        SchemaWith(pEl))

    Dim pIl As Object
    Set pIl = CreateObject("Scripting.Dictionary")
    Set pIl("path") = Prop("string", "Markdown file to import")
    Set pIl("dryRun") = Prop("boolean", "Parse only, do not write (default false)")
    tools.Add MakeTool("import_lessons_md", _
        "Import lessons from a Markdown file: a line '## [id] scenario' plus trap:/do: lines.", _
        SchemaWith(pIl, Req("path")))

    ' ---- extension library (扩展库): declarative tool packs ----
    Dim pXl As Object
    Set pXl = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("ext_list", _
        "List installed extension packs with enabled state, tool count and file path.", _
        SchemaWith(pXl))

    Dim pXe As Object
    Set pXe = CreateObject("Scripting.Dictionary")
    Set pXe("name") = Prop("string", "Extension pack name")
    tools.Add MakeTool("ext_enable", "Enable an extension pack; its tools appear in tools/list.", SchemaWith(pXe))

    Dim pXd As Object
    Set pXd = CreateObject("Scripting.Dictionary")
    Set pXd("name") = Prop("string", "Extension pack name")
    tools.Add MakeTool("ext_disable", "Disable an extension pack; its tools disappear from tools/list.", SchemaWith(pXd))

    Dim pXr As Object
    Set pXr = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("ext_reload", "Rescan mcp_extensions and report pack / enabled-tool counts.", SchemaWith(pXr))

    Dim pXi As Object
    Set pXi = CreateObject("Scripting.Dictionary")
    Set pXi("path") = Prop("string", "Extension JSON file to install")
    tools.Add MakeTool("ext_install", "Install an extension pack by copying its JSON into mcp_extensions.", SchemaWith(pXi, Req("path")))

    Dim pXu As Object
    Set pXu = CreateObject("Scripting.Dictionary")
    Set pXu("name") = Prop("string", "Extension pack name")
    tools.Add MakeTool("ext_uninstall", "Remove an extension pack (file renamed to .removed).", SchemaWith(pXu))

    ' merge the enabled extension tools into the advertised list
    Dim extSpecs As Collection, ei As Long, eItem As Object
    Dim scRaw As Variant, scIsObj As Boolean, scUse As Object
    On Error Resume Next
    Set extSpecs = Extensions.ToolSpecs()
    On Error GoTo 0
    If Not extSpecs Is Nothing Then
        For ei = 1 To extSpecs.Count
            Set eItem = extSpecs(ei)
            Set scUse = Nothing
            If JsonGet(eItem("spec"), "inputSchema", scRaw, scIsObj) Then
                If scIsObj Then
                    If TypeName(scRaw) = "Dictionary" Then Set scUse = scRaw
                End If
            End If
            If scUse Is Nothing Then Set scUse = SchemaWith(Nothing)
            tools.Add MakeTool(CStr(eItem("name")), CStr(eItem("description")), scUse)
        Next
    End If

    ' ---- VBA6 debug-engine probe: the real execution entry point ----
    Dim pVb As Object
    Set pVb = CreateObject("Scripting.Dictionary")
    Set pVb("run_break_probes") = Prop("boolean", "Also call the break-only probes (EbGetCallstackCount/Function, EbGetCurrentFunction, EbGetDebWndCallstackIndex). Requires a break context.")
    Set pVb("force") = Prop("boolean", "Run break probes even when EbMode reports no break context (may be unsafe)")
    Set pVb("ident") = Prop("string", "Optional identifier to test with EbIsValidIdent (design time)")
    Set pVb("exec_probe") = Prop("boolean", "Run the Immediate-window executor probe (EbExecuteLine) instead of the standard set. RE round 5: this export (ordinal 147, RVA 0x127065) IS the Immediate-window line executor.")
    Set pVb("exec_line") = Prop("string", "Line to execute, e.g. 1+1 or Debug.Print ... (Immediate-window syntax)")
    Set pVb("exec_a2") = Prop("integer", "arg2 (semantics unresolved; values 0/1 are the candidates)")
    Set pVb("exec_a3") = Prop("integer", "arg3 (unresolved)")
    Set pVb("exec_a4") = Prop("integer", "arg4 (mode flag: the epilogue uses setz al, so 0 vs 1 differ)")
    Set pVb("exec_scope") = Prop("boolean", "Set the eval-scope bit (WORD[VBA6+0x17D076] |= 0x100) around the call, the way the IDE's own immediate-window path does, then restore it. This is the cheapest test of whether locals become visible at a breakpoint.")
    Set pVb("exec_dump_ctx") = Prop("boolean", "Pass arg2 a 64-byte buffer and dump the first 32 bytes afterwards. WARNING: measured to crash the IDE (AV in VBA6.dll) - arg2 is dereferenced as a structured pointer, so leave it 0 unless the object type is known.")
    Set pVb("exec_set_frame") = Prop("boolean", "F-2 (RE round 9): make the breakpoint's frame the current evaluation frame (write dword_global VBA6+0x17D0A4 from the engine+0x2F0 chain's first kind==1 node payload) around the call, then restore. Expected to make locals visible at a breakpoint.")
    Set pVb("frame_probe") = Prop("boolean", "F-1: read-only two-state observation - EbMode, dword_global VBA6+0x17D0A4 (evaluation frame), and the chain at engine+0x2F0 with each node's kind/payload. Writes nothing.")
    Set pVb("frame_set") = Prop("boolean", "F-3: call the engine's own frame setter sub_FBC3163 (RVA 0x133163) through an executable thunk, with frame_set_a1/frame_set_a2 (default 0/1 = find the named frame). This is the authoritative way to bind the evaluation frame.")
    Set pVb("frame_set_a1") = Prop("integer", "F-3 arg1: frame index (default 0)")
    Set pVb("frame_set_a2") = Prop("integer", "F-3 arg2: 1 = walk to the first kind==1 frame (default 1)")
    Set pVb("rva_call") = Prop("string", "Call an INTERNAL function by RVA; optional vb6: prefix targets the HOST image VB6.EXE instead of VBA6.DLL. Call a VBA6 function by RVA (hex, e.g. 0x23EA) through an executable thunk, with rva_argc/rva_a1/rva_a2. Unsafe by nature: the callee must be __stdcall with argc args and clean its own stack.")
    Set pVb("rva_argc") = Prop("integer", "Argument count 0..2 (must match the callee's ret N)")
    Set pVb("rva_a1") = Prop("integer", "Argument 1")
    Set pVb("rva_a2") = Prop("integer", "Argument 2")
    Set pVb("worker_call") = Prop("boolean", "DECISIVE (RE round 15): call the real worker sub_FB7317A (RVA 0xE317A) the way the IDE does - __thiscall ecx=engine, 8 stack args, retn 20h. Can crash the IDE if an argument is wrong.")
    Set pVb("worker_text") = Prop("string", "Line to evaluate, e.g. an OutputDebugStringA call that prints CStr(n)")
    Set pVb("worker_bstr") = Prop("boolean", "Pass the text as a BSTR (IDE form, default) instead of an ANSI buffer (EbExecuteLine form)")
    Set pVb("worker_set_bit") = Prop("boolean", "Also set WORD[VBA6+0x17D076] |= 0x100 around the call (the IDE path does this)")
    Set pVb("worker_p1") = Prop("integer", "worker stack arg p1 (IDE: 1)")
    Set pVb("worker_p2") = Prop("integer", "worker stack arg p2 (IDE: 1)")
    Set pVb("worker_p3") = Prop("integer", "worker stack arg p3 (IDE: 0; step path: LCID)")
    Set pVb("worker_p4") = Prop("integer", "worker stack arg p4 (IDE: 0; step path: MODULE HANDLE - the prime suspect)")
    Set pVb("worker_p5") = Prop("integer", "worker stack arg p5 (IDE: 1)")
    Set pVb("worker_p6") = Prop("integer", "worker stack arg p6 (IDE: 1)")
    Set pVb("worker_p7") = Prop("integer", "worker stack arg p7 (IDE: 0)")
    Set pVb("worker_p2_engine") = Prop("boolean", "Pass p2 = the engine singleton instead of the IDE's literal 1. RE round-16 §5 note 2: with p1=0 the worker may dereference p2 as an object, and 1 then crashes.")
    Set pVb("mem_dump") = Prop("boolean", "READ-ONLY memory dump for the RE round-17 P-series (locate the ExFrame stack top / check frame offsets). Safe: refuses addresses outside user space.")
    Set pVb("mem_rva") = Prop("string", "Address as an RVA (hex, e.g. 0x17D000). Default image is VBA6.DLL; prefix vb6+ to address the HOST image VB6.EXE (e.g. vb6+0x1234) or vba6+ to be explicit. Used when mem_addr is 0")
    Set pVb("mem_addr") = Prop("string", "Absolute address to read (hex, e.g. 0x01234567); takes precedence over mem_rva")
    ' RE round-23/24 additions (probe-only; the two are independent of each other)
    Set pVb("hosted_expr_srv") = Prop("boolean", "NEW DIRECTION: instantiate VBA6's HOSTED expression service via CreateHostedExprSrv (one of the 317 exports the IDE never imports), then report the object pointer it returns and the first 10 slots of its vtable. Read-mostly: it allocates a 404-byte engine object; nothing is patched.")
    Set pVb("wnd_probe") = Prop("boolean", "READ-ONLY, in-process: enumerate the DIRECT CHILDREN (depth exactly 2) of this process's TOP-LEVEL windows and classify both the window procedure and the module that registered the window class (vb6 / vba6 / other). The top-level windows themselves are NEVER listed and grandchildren or deeper are NEVER listed; there are no visibility, owner, rect or style filters (source: Vba6Probe.bas WindowProbeJson). Cross-process GetWindowLong/GCL_HMODULE both return 0, so this probe only works from inside the IDE. WARNING: this is NOT a complete window enumeration - it saw only 27 of the 76 windows that find_gui returned (about 36 percent), so any does-not-exist conclusion about a window or a dialog drawn from wnd_probe is invalid. STRUCTURAL BLIND SPOT: a dialog is normally a TOP-LEVEL #32770 window, which wnd_probe structurally cannot see - use find_gui (all top-level windows plus the whole recursive subtree) to find dialogs or any complete set.")
    Set pVb("wnd_max") = Prop("integer", "Maximum windows to report for wnd_probe (default 200, cap 1500)")
    Set pVb("mem_scan") = Prop("boolean", "READ-ONLY byte-wise dword scanner over a live range: reports every offset whose dword lies in [scan_lo, scan_hi], with its alignment. Used to replicate RE round-32's whole-image scan of VB6.EXE for absolute references into the VBA6 engine globals.")
    Set pVb("scan_addr") = Prop("string", "Start address for mem_scan (hex, e.g. 0x400000)")
    Set pVb("scan_len") = Prop("integer", "Bytes to scan for mem_scan (max 1048576)")
    Set pVb("scan_lo") = Prop("integer", "Low bound of the dword value range to match (must be non-zero)")
    Set pVb("scan_hi") = Prop("integer", "High bound of the dword value range to match")
    Set pVb("scan_max") = Prop("integer", "Maximum hits to return (default 64, cap 512)")
    Set pVb("export_rvas") = Prop("boolean", "READ-ONLY: resolve the exports the RE names BY NAME through the real PE export table and report each one's true RVA, so any 'function X is at RVA 0xNNNN' claim can be checked in one call.")
    Set pVb("engine_items") = Prop("boolean", "READ-ONLY: call EbGetStatusOfItems(buf,72) the way VB6.EXE's command dispatcher does and return the engine's 72-item capability table (hex + list of enabled item ids). No writes.")
    Set pVb("cs_fn") = Prop("integer", "Callback-free single-frame probe: call EbGetCallstackFunction(frame, &outA, &outB, &outC, &arg5) with the authoritative IDE prototype (3 out-BSTRs zeroed, arg5 a zeroed writable dword, rc=0 means success). frame must be 0..count-1. Omit or -1 to skip.")
    Set pVb("mem_len") = Prop("integer", "Bytes to read, multiple of 4, max 512 (default 64)")
    ' ---- project change journal (P0-1: IDE event sinks) ----
    Dim pChg As Object
    Set pChg = CreateObject("Scripting.Dictionary")
    Set pChg("since") = Prop("integer", "Return only changes with rev > since (0 = everything retained)")
    Set pChg("clear") = Prop("boolean", "Reset the dropped-events counter after reading")
    tools.Add MakeTool("project_changes", _
        "Poll the IDE change journal instead of re-reading the whole model: project/component/control add, remove and rename plus component reload, each with a monotonically increasing revision. Keep the last revision and pass it as since on the next call.", _
        SchemaWith(pChg))

    ' ---- call stack (G1): per-frame module + procedure, straight from the engine ----
    Dim pCs As Object
    Set pCs = CreateObject("Scripting.Dictionary")
    Set pCs("max_frames") = Prop("integer", "Maximum frames to return (default 32, cap 128)")
    tools.Add MakeTool("get_callstack", _
        "Read the engine's OWN call stack while the program is stopped at a breakpoint: for EVERY frame the project, the MODULE and the PROCEDURE name, indexed 0 (innermost) to count-1 (outermost). Closes the G1 gap (per-frame procedure name) with the same EbGetCallstackFunction call that the IDE's own Call Stack window uses. Requires break mode; otherwise returns broken=false with a note. Implementation note: the call uses the corrected machine-code thunk prototype with the frame index CLAMPED to 0..count-1; the naive ByRef-String declaration of EbGetCallstackFunction is a different, IDE-killing path and is never used.", _
        SchemaWith(pCs))

    tools.Add MakeTool("vba6_probe", _
        "Call the VBA6.DLL debug engine from inside the IDE process: EbMode / EbGetStepMode / callstack / current function / EbIsValidIdent, or (exec_probe=true) the Immediate-window executor EbExecuteLine. Read-only probes run by default; break-only probes and the executor are gated.", _
        SchemaWith(pVb))

    ' ---- project metadata / version resource (text-level .vbp) ----
    Dim pPm As Object
    Set pPm = CreateObject("Scripting.Dictionary")
    Dim pPm_action As Object
    Set pPm_action = Prop("string", "get (read current values) or set (apply the supplied fields)")
    Set pPm_action("enum") = EnumList("get", "set")
    Set pPm("action") = pPm_action
    Set pPm("path") = Prop("string", "Project .vbp path; defaults to the active project when the IDE is connected")
    Set pPm("major") = Prop("string", "MajorVer")
    Set pPm("minor") = Prop("string", "MinorVer")
    Set pPm("revision") = Prop("string", "RevisionVer")
    Set pPm("auto_increment") = Prop("boolean", "AutoIncrementVer (1 = bump on every build)")
    Set pPm("company_name") = Prop("string", "CompanyName")
    Set pPm("file_description") = Prop("string", "FileDescription")
    Set pPm("product_name") = Prop("string", "ProductName")
    Set pPm("legal_copyright") = Prop("string", "LegalCopyright")
    Set pPm("legal_trademark") = Prop("string", "LegalTrademark")
    Set pPm("comments") = Prop("string", "Comments")
    Set pPm("title") = Prop("string", "Title")
    Set pPm("exe_name") = Prop("string", "ExeName32 (output file name)")
    Set pPm("cond_comp") = Prop("string", "CondComp (conditional compilation arguments, e.g. TESTING=1)")
    tools.Add MakeTool("project_meta", _
        "Read or write the .vbp version resource and build metadata (MajorVer/MinorVer/RevisionVer/AutoIncrementVer/CompanyName/FileDescription/ProductName/legal notices/Comments/Title/ExeName32/CondComp). These keys are not exposed by the VBIDE object model, so they are edited as text: component and reference lines are never touched and a .bak is kept.", _
        SchemaWith(pPm, Req("action")))

    ' ---- application manifest (UAC / DPI) ----
    Dim pMf As Object
    Set pMf = CreateObject("Scripting.Dictionary")
    Dim pMf_action As Object
    Set pMf_action = Prop("string", "read (show the existing manifest) or write (generate/overwrite it)")
    Set pMf_action("enum") = EnumList("read", "write")
    Set pMf("action") = pMf_action
    Set pMf("path") = Prop("string", "The .exe/.dll the manifest belongs to (the file <path>.manifest is used)")
    Set pMf("level") = Prop("string", "UAC level: asInvoker / highestAvailable / requireAdministrator")
    Set pMf("dpi_aware") = Prop("boolean", "Emit dpiAware=true (default true)")
    Set pMf("dpi_awareness") = Prop("string", "Optional 2016 setting: permonitorv2 / permonitor / system")
    Set pMf("ui_access") = Prop("boolean", "requestedExecutionLevel uiAccess (default false)")
    Set pMf("assembly_name") = Prop("string", "assemblyIdentity name (default vb6mcp.app)")
    tools.Add MakeTool("app_manifest", _
        "Generate or inspect an application manifest next to an exe/dll (UAC requestedExecutionLevel plus DPI awareness). Writes <path>.manifest, which Windows reads automatically for an unembedded external manifest.", _
        SchemaWith(pMf, Req("action", "path")))

    ' ---- missing-reference check ----
    Dim pRc As Object
    Set pRc = CreateObject("Scripting.Dictionary")
    Set pRc("include_resolved") = Prop("boolean", "Also list references that resolve fine (default false = problems only)")
    tools.Add MakeTool("check_references", _
        "List the active project's references and flag any whose file is missing on disk - the classic missing-reference failure of older VB6 projects. Built-in references (VBA/VBRUN/VB) are always treated as resolved because VB6 reports a synthetic path for them.", _
        SchemaWith(pRc))

    ' ---- project group (.vbg) ----
    Dim pVg As Object
    Set pVg = CreateObject("Scripting.Dictionary")
    Dim pVg_action As Object
    Set pVg_action = Prop("string", "list / add / remove a project in a .vbg project group")
    Set pVg_action("enum") = EnumList("list", "add", "remove")
    Set pVg("action") = pVg_action
    Set pVg("path") = Prop("string", "The .vbg project-group file")
    Set pVg("project") = Prop("string", "The .vbp to add or remove (action=add/remove)")
    tools.Add MakeTool("vbg_group", _
        "Read or edit a .vbg project group: list its projects (with existence flags) or add/remove one. Text level, so it works whether or not the IDE has the group open.", _
        SchemaWith(pVg, Req("action", "path")))

    ' ---- project-level snapshot / rollback ----
    Dim pSn As Object
    Set pSn = CreateObject("Scripting.Dictionary")
    Dim pSn_action As Object
    Set pSn_action = Prop("string", "create (snapshot the .vbp plus all component files), restore (roll back), or list")
    Set pSn_action("enum") = EnumList("create", "restore", "list")
    Set pSn("action") = pSn_action
    Set pSn("path") = Prop("string", "The .vbp to snapshot. Required for action=create and action=list; NOT used by action=restore")
    Set pSn("dir") = Prop("string", "Snapshot directory. Used by action=create as an override (default <project>\_vb6mcp_snapshots\<timestamp>); REQUIRED by action=restore (the directory to restore from)")
    tools.Add MakeTool("project_snapshot", _
        "Project-level checkpoint and rollback: copy the .vbp plus every component file into a timestamped folder with a snapshot.json manifest (path, sha256, size per file), then restore any snapshot later. Every overwritten file is backed up as .bak. This is the project-wide counterpart of the component-level snapshot/verify tools.", _
        SchemaWith(pSn, Req("action")))

    ' ---- delivery packaging ----
    Dim pPk As Object
    Set pPk = CreateObject("Scripting.Dictionary")
    Set pPk("path") = Prop("string", "The .vbp to package")
    Set pPk("exe_path") = Prop("string", "Built exe/dll; defaults to the ExeName32 of the project")
    Set pPk("out_dir") = Prop("string", "Output folder; defaults to <project>\_vb6mcp_dist")
    Set pPk("level") = Prop("string", "UAC level for the generated manifest (asInvoker / highestAvailable / requireAdministrator)")
    tools.Add MakeTool("package_app", _
        "Assemble a delivery folder for a project: the built exe, an external manifest (UAC + per-monitor DPI), dependencies.txt built from the .vbp Reference= lines (with sha256 and MISSING flags), an install.bat that regsvr32s the OCX/DLL dependencies, and a README. Pure file level, so it works while the IDE is closed.", _
        SchemaWith(pPk, Req("path")))

    ' ---- read-only VCS ----
    Dim pGt As Object
    Set pGt = CreateObject("Scripting.Dictionary")
    Dim pGt_action As Object
    Set pGt_action = Prop("string", "status (porcelain + branch), diff (--stat, optional file filter) or log (oneline)")
    Set pGt_action("enum") = EnumList("status", "diff", "log")
    Set pGt("action") = pGt_action
    Set pGt("repo") = Prop("string", "Repository directory; defaults to the active project folder")
    Set pGt("file") = Prop("string", "Optional path filter for action=diff")
    Set pGt("max") = Prop("integer", "Number of commits for action=log (default 20)")
    Set pGt("exe") = Prop("string", "Path to git.exe (default: resolved from PATH)")
    tools.Add MakeTool("git", _
        "Read-only Git integration: status (porcelain v1 plus branch header), diff (--stat, optionally for one file) and log (oneline). Runs git.exe and returns its captured stdout, so an agent can see what changed before editing or committing. No write commands are exposed.", _
        SchemaWith(pGt, Req("action")))

    ' ---- offline VB tooling (VbFmt.bas) ----
    Dim pFc As Object
    Set pFc = CreateObject("Scripting.Dictionary")
    Set pFc("path") = Prop("string", "VB source file to re-indent")
    Set pFc("indent") = Prop("string", "Indent unit (default four spaces; use two spaces or a tab if you prefer)")
    Set pFc("apply") = Prop("boolean", "false (default) = report only; true = rewrite the file (a .bak is not needed: only leading whitespace can change and that is verified before writing)")
    tools.Add MakeTool("format_code", _
        "Re-indent a VB6 source file from its block structure (Sub/Function/If/For/Do/Select/With/Type/Enum plus #If and Case/Else). It rewrites leading whitespace ONLY and proves it: if the stripped text would differ on any line it refuses to write. Strings and comments are masked before keyword detection, so keywords inside literals are never mistaken for blocks.", _
        SchemaWith(pFc, Req("path")))

    Dim pRn As Object
    Set pRn = CreateObject("Scripting.Dictionary")
    Set pRn("path") = Prop("string", "The .vbp whose component files are searched")
    Set pRn("old_name") = Prop("string", "Identifier to rename")
    Set pRn("new_name") = Prop("string", "New identifier (must be a valid VB identifier)")
    Set pRn("dry_run") = Prop("boolean", "true (default) = only report the occurrence counts")
    Set pRn("skip_attributes") = Prop("boolean", "Skip Attribute lines (default true) - module names need rename_component instead")
    tools.Add MakeTool("rename_symbol", _
        "Rename an identifier (variable, procedure, type) across every component of a project, with a per-file occurrence count and sha256 before/after. String literals and comments are skipped by a lexical scan, and Attribute lines are skipped by default, so a component rename is routed to rename_component instead. dry_run defaults to true; each rewritten file keeps a .bak.", _
        SchemaWith(pRn, Req("path", "old_name", "new_name")))

    ' ---- native IDE breakpoints (menu path; no source modification) ----
    Dim pNb As Object
    Set pNb = CreateObject("Scripting.Dictionary")
    Dim pNb_action As Object
    Set pNb_action = Prop("string", "toggle (Debug menu Toggle Breakpoint at component/line) or clear_all (Clear All Breakpoints)")
    Set pNb_action("enum") = EnumList("toggle", "clear_all")
    Set pNb("action") = pNb_action
    Set pNb("component") = Prop("string", "Component whose code pane holds the line (action=toggle)")
    Set pNb("line") = Prop("integer", "1-based IDE line to place the caret on before toggling (action=toggle)")
    tools.Add MakeTool("native_breakpoint", _
        "Toggle a REAL IDE breakpoint (or clear all) through the Debug menu without touching any source file - unlike set_breakpoint, which injects Debug.Assert False into the code. Places the code pane caret first. Captions are matched for English and Chinese IDEs; note that the CommandBars model is suspended while the program runs, so use this in design or break state.", _
        SchemaWith(pNb, Req("action")))

    Set root("tools") = tools
    Dim pWss As Object
    Set pWss = CreateObject("Scripting.Dictionary")
    Set pWss("repo") = Prop("string", "Repository directory for the git half (default: the active project folder)")
    tools.Add MakeTool("workspace_status", _
        "Read-only: list the IDE components that are UNSAVED and the git working tree (branch/ahead/behind/files/diffStat) side by side. Runs only read-only git subcommands and never writes.", _
        SchemaWith(pWss))

    Dim pVs As Object
    Set pVs = CreateObject("Scripting.Dictionary")
    Set pVs("components") = Prop("array", "Optional component names to check (default: every component of the active project)")
    Set pVs("path") = Prop("string", "Optional disk path override")
    tools.Add MakeTool("verify_scope", _
        "Read-only: for a set of components report {component, ideEqualsDisk, lintErrors, lintWarnings, firstDiagnostic} with a pass/fail summary; only failing components are listed in detail.", _
        SchemaWith(pVs))

    tools.Add MakeTool("todo_status", _
        "Read-only: machine-readable todo derived from _bus/tasks/P1 and _bus/reports/P1 (status + title) plus counts; useful for a quick progress snapshot.", _
        SchemaWith(Nothing))

    ' ---- contract v1.3.0 (TASK-P1-030): machine-readable version identity ----
    tools.Add MakeTool("mcp_info", _
        "Read-only version stamp for this server: {ok, product, contract, toolCount, build:{sha256, builtAt}, breaking[], migrations[], changelog[]}. build.sha256 is the SHA-256 of the DLL file on disk, read at call time (never a compile-time literal), so an independent Get-FileHash can be compared with it; builtAt is that file's LastWriteTime. breaking[] lists the contract changes that alter how existing calls must be made. Call this once after initialize to learn which contract the server implements.", _
        SchemaWith(Nothing))

    ' ---- contract v1.6.0 (TASK-P1-076): cross-component replace + component undo ----
    Dim pRif As Object
    Set pRif = CreateObject("Scripting.Dictionary")
    Set pRif("path") = Prop("string", "The .vbp of the OPEN project, or its directory. Any other path is refused: this tool never writes a project the add-in does not own")
    Set pRif("query") = Prop("string", "Literal text, or an ECMAScript-style regex when regex=true")
    Set pRif("replacement") = Prop("string", "Replacement text; with regex=true, $1..$9 refer to capture groups")
    Set pRif("regex") = Prop("boolean", "Treat query as a regular expression (default false = literal)")
    Set pRif("dry_run") = Prop("boolean", "Default TRUE: plan only, write nothing. Set false to apply.")
    Set pRif("expected_shas") = Prop("array", "Optional per-component guard: [{file, sha256}] where file is a component name or its file path; a mismatch SKIPS that component and reports it (never written)")
    Set pRif("validate") = Prop("string", "off|warn|strict (default warn) - lint gate applied before each write, as in the other write tools")
    Set pRif("max_files") = Prop("integer", "Optional cap on components scanned (default 200, cap 1000) so a large project cannot hang the IDE")
    tools.Add MakeTool("replace_in_files", _
        "Cross-component literal/regex replace over the components of the OPEN project. Defaults to dry_run=true: nothing is written and the receipt lists every matching component with its hit count and the shas it WOULD produce. With dry_run=false each component is rewritten through one minimal text transaction (per-component atomic) and the receipt reports per-component shaBefore/shaAfter plus a summary. expected_shas[] guards each component: a mismatch makes that component SKIPPED and reported, never written. Failure semantics are honest: on the first failure the tool STOPS, ROLLS BACK the components it already applied to their pre-run text (rolledBack=true) or reports the exact applied set when a rollback is not possible - it never claims atomicity across components.", _
        SchemaWith(pRif, Req("path", "query", "replacement")))

    Dim pUnd As Object
    Set pUnd = CreateObject("Scripting.Dictionary")
    Set pUnd("component") = Prop("string", "Component name; omit to use the component of the most recent snapshot")
    Dim pUnd_action As Object
    Set pUnd_action = Prop("string", "list = show the available snapshots; restore (default) = restore one")
    Set pUnd_action("enum") = EnumList("list", "restore")
    Set pUnd("action") = pUnd_action
    Set pUnd("ts") = Prop("string", "Optional exact snapshot timestamp (see action=list); omit for the most recent")
    Set pUnd("to_disk") = Prop("boolean", "Default FALSE: restore changes the component's PANE text only, the file on disk is NEVER touched by default. true is not supported yet and is refused with reason=to_disk_not_supported - use save_project to flush the project.")
    tools.Add MakeTool("undo_edit", _
        "Component-level undo of the LAST edit, from snapshots taken before every text write (the pre-write PANE text is stored as UTF-8 in <project>_undo<component>.<timestamp>.txt). Explicit and single-step: it never rolls back automatically and never touches a file on disk. action=list shows the available snapshots; action=restore (default) restores one through the same minimal-write path and reports the sha before/after plus diskUntouched.", _
        SchemaWith(pUnd))

    ' ---- contract v1.9.0 (TASK-P1-164 A / TASK-P1-168 B): foreground safety + immediate execution ----
    Dim pImE As Object
    Set pImE = CreateObject("Scripting.Dictionary")
    Set pImE("expression") = Prop("string", "Expression to evaluate in the ACTIVE project's immediate context, e.g. 1/4 or UCase$(""ab""). It is evaluated exactly ONCE.")
    Set pImE("timeout_ms") = Prop("integer", "Gate-age threshold in ms; if this call never returns, the ARMED release frees the single-flight gate on the next tool call (default 5000). It is NOT a hard timeout - the add-in cannot preempt a statement.")
    tools.Add MakeTool("immediate_eval", _
        "Evaluate an EXPRESSION in the VB6 IDE's own immediate context and return its VALUE, with the caller NOT having to write Open/Print/Close: {ok, value, typeName, output, elapsedMs}. The value is carried out through a private temp file the tool reads back, so no Immediate-window pane read is involved. An unparseable sink is reported as {ok:false, reason:""value_not_parsed""} together with the RAW output - nothing is guessed. A statement that RAISES still blocks the IDE (VB6 refuses On Error on this channel with ret=40036, measured), so timeout_ms is a gate-age threshold rather than a hard timeout; ide_state{release_inflight:true} forces the gate open afterwards.", _
        SchemaWith(pImE, Req("expression")))

    Dim pImR As Object
    Set pImR = CreateObject("Scripting.Dictionary")
    Set pImR("code") = Prop("string", "One or more statements to run, e.g. Debug.Print 1+1 or x = 5: Debug.Print x. A CR/LF in the input is joined with a colon and a space; each statement-boundary Debug.Print is redirected into the sink.")
    Set pImR("timeout_ms") = Prop("integer", "Gate-age threshold in ms; if this call never returns, the ARMED release frees the single-flight gate on the next tool call (default 5000). It is NOT a hard timeout.")
    tools.Add MakeTool("immediate_run", _
        "RUN VB statements in the VB6 IDE's own immediate context and return what they PRINTED, with the caller NOT having to write Open/Print/Close: {ok, output, outputTrimmed, outputLines, debugPrintRewrites, elapsedMs}. Each statement-boundary Debug.Print is redirected into a private temp file, and only the text produced by THIS call is returned. The ? shorthand and _ line continuations are not supported. A statement that RAISES still blocks the IDE - see immediate_eval; timeout_ms arms a self-healing release and ide_state{release_inflight:true} forces the gate open.", _
        SchemaWith(pImR, Req("code")))

    Dim pQm As Object
    Set pQm = CreateObject("Scripting.Dictionary")
    Set pQm("on") = Prop("boolean", "true = turn quiet mode ON (the add-in stops taking the user's foreground and every receipt carries a quiet{} block with an advisory); false = turn it OFF and restore the previous behaviour byte for byte.")
    Set pQm("ttl_ms") = Prop("integer", "Optional lifetime in ms (default 600000 = 10 minutes). The mode expires by itself once the TTL passes; the TTL is re-evaluated on every read.")
    tools.Add MakeTool("ide_quiet_mode", _
        "Turn the SESSION-LEVEL quiet mode on or off. While ON the add-in does not call Window.SetFocus for control_window{action:""activate""} (that receipt reports quietModeSuppressedFocus:true plus the action that was still performed) and it gives the user's foreground back after a menu-driven action (native_breakpoint); EVERY tools/call receipt then carries quiet{quietMode,since,ttlMs,quietExpiresAt,suppressedFocusCount,advisory}. The advisory asks the human not to type into this IDE while an agent drives it, because keystrokes aimed at another window can land in the IDE when it is raised. quiet OFF is byte-for-byte the previous behaviour, and foregroundReadable=false means the foreground handle could not be READ (it is 0 here) - never read that as ""the focus was left alone"". This tool is read-only: it changes only the add-in's own state, so it is still served while a write is in flight.", _
        SchemaWith(pQm, Req("on")))

    Set BuildToolsListResult = root
End Function

' =====================================================================
' contract v1.3.0: build / version identity (TASK-P1-030)
' =====================================================================

Public Function ProductVersion() As String
    ProductVersion = VB6MCP_PRODUCT_VERSION
End Function

Public Function ContractVersion() As String
    ContractVersion = VB6MCP_CONTRACT_VERSION
End Function

' The DLL that is currently loaded. App.hInstance is this ActiveX DLL's module
' handle, so GetModuleFileNameW yields the very file the hash below describes.
Private Function OwnDllPath() As String
    On Error Resume Next
    Dim buf As String
    buf = String$(1024, Chr$(0))
    Dim n As Long
    n = OwnModuleFileNameW(App.hInstance, StrPtr(buf), 1024)
    If n > 0 Then OwnDllPath = Left$(buf, n)
    If Len(OwnDllPath) = 0 Then OwnDllPath = App.Path & "\" & App.EXEName & ".dll"
    Err.Clear
End Function

' SHA-256 of VB6MCP.dll AS IT IS ON DISK, read at call time - never a
' compile-time literal, so an independent Get-FileHash must agree with it.
Public Function BuildSha256() As String
    On Error Resume Next
    Dim p As String
    p = OwnDllPath()
    If Len(p) = 0 Then Exit Function
    BuildSha256 = Encoding.Sha256HexFile(p)
End Function

Public Function BuildBuiltAt() As String
    On Error Resume Next
    Dim p As String
    p = OwnDllPath()
    If Len(p) = 0 Then Exit Function
    BuildBuiltAt = Format$(FileDateTime(p), "yyyy-mm-dd hh:nn:ss")
End Function

' The exact file the two stamps above describe (TASK-P1-036 auditability).
Public Function BuildPath() As String
    On Error Resume Next
    BuildPath = OwnDllPath()
End Function

' Tool count straight from the tool table (never a hardcoded literal), so it
' always agrees with tools/list.
Public Function ToolCount() As Long
    On Error Resume Next
    Dim root As Object, tools As Object
    Set root = BuildToolsListResult()
    If root Is Nothing Then Exit Function
    Set tools = root("tools")
    If tools Is Nothing Then Exit Function
    ToolCount = tools.Count
End Function

' Contract v1.3.0 mechanism 1: the fixed banner an agent reads before calling.
' The FIRST line is exactly (two spaces between the groups):
'   PRODUCT vb6-mcp <semver>  CONTRACT <contract>  TOOLS <n>  BUILT <sha256 first 8>
' immediately followed by the BREAKING list, then the existing lessons block.
Public Function BuildInstructions() As String
    Dim s As String
    s = "PRODUCT " & VB6MCP_PRODUCT & " " & VB6MCP_PRODUCT_VERSION & "  CONTRACT " & VB6MCP_CONTRACT_VERSION & "  TOOLS " & CStr(ToolCount()) & "  BUILT " & Left$(BuildSha256(), 8) & vbCrLf
    s = s & "BREAKING (read before calling):" & vbCrLf
    s = s & "- get_component_code: coords default disk -> ide (pass coords=""disk"" for the old text/numbering); response is a JSON envelope, text in result.text" & vbCrLf
    s = s & "- replace_code: out-of-range/continuation now REJECTED (range_beyond_end / inside_continuation); deletedLines=0 is no longer a success signal" & vbCrLf
    s = s & "- write-class tools: validate=off|warn|strict; edit-side failures carry ok:false + reason" & vbCrLf
    s = s & "- write_file now truncates; build/make_project/verify: judge by artifactChanged/verdict (compiled-no-artifact-change = NOT built)" & vbCrLf
    s = s & "- v1.1.0 added ide_state/current_position/list_breakpoints/find_gui/ui_text; v1.2.0 added workspace_status/verify_scope/todo_status" & vbCrLf
    s = s & "- v1.7.0 snapshot param (ADDITIVE, non-breaking): replace_code / apply_edits / edit_by_anchor accept snapshot=true (on/yes/1) to guarantee exactly ONE pre-edit pane snapshot that undo_edit can restore; default off (set / apply_edits / edit_by_anchor already snapshot by default)" & vbCrLf
    Dim les As String
    On Error Resume Next
    les = Lessons.Instructions()
    Err.Clear
    On Error GoTo 0
    If Len(les) > 0 Then s = s & vbCrLf & les
    BuildInstructions = s
End Function

' Contract v1.3.0 mechanism 2: the version stamp injected into every tools/call
' text payload. It is a key INSIDE the tool's JSON body - never a top-level
' JSON-RPC result key - so tools/call result keys stay exactly {content,isError}.
Private Function McpQ(ByVal s As String) As String
    McpQ = Chr$(34) & s & Chr$(34)
End Function

Public Function McpStampJson() As String
    ' TASK-P1-037 (contract v1.3.1 H2): the receipt also says WHICH IDE answered -
    ' pid/project/port - alongside the mcp contract stamp. Both go into the tool
    ' BODY, never into a top-level JSON-RPC result key.
    McpStampJson = McpQ("mcp") & ":{" & McpQ("product") & ":" & McpQ(VB6MCP_PRODUCT_VERSION) & "," & McpQ("contract") & ":" & McpQ(VB6MCP_CONTRACT_VERSION) & "}," & _
        McpQ("ide") & ":{" & McpQ("pid") & ":" & CStr(GetCurrentProcessId()) & "," & McpQ("project") & ":" & McpQ(ProjectNameSafe()) & "," & McpQ("port") & ":" & CStr(ServerPort()) & "}"
End Function

' contract v1.3.1 H2: the port this instance actually listens on (0 = not listening).
Public Function ServerPort() As Long
    On Error Resume Next
    ServerPort = 0
    If m_server Is Nothing Then Exit Function
    ServerPort = m_server.Port
End Function

' Insert the stamp as the FIRST key of a JSON *object* body. A body that is not
' a JSON object is returned unchanged (export_lessons returns Markdown by
' design), so a text payload is never corrupted.
Private Function StampMcp(ByVal body As String) As String
    If Len(body) = 0 Then
        StampMcp = body
        Exit Function
    End If
    If Left$(body, 1) <> "{" Then
        StampMcp = body
        Exit Function
    End If
    ' TASK-P1-164 (A): while ide_quiet_mode is ON every receipt carries the quiet{} block
    ' (and therefore the advisory). While it is OFF QuietStamp() returns "" and the string
    ' built below is character-for-character the pre-v1.9.0 one.
    Dim q As String
    q = QuietStamp()
    If Len(body) = 2 Then
        ' "{}" must not gain a trailing comma
        StampMcp = "{" & McpStampJson() & q & "}"
    Else
        StampMcp = "{" & McpStampJson() & q & "," & Mid$(body, 2)
    End If
End Function

' TASK-P1-164 (A): the leading-comma fragment of the quiet{} block, or "" while quiet
' mode is OFF. Kept out of every tool body so no tool has to know about it.
Private Function QuietStamp() As String
    If Not QuietActive() Then
        QuietStamp = ""
        Exit Function
    End If
    QuietStamp = "," & Chr$(34) & "quiet" & Chr$(34) & ":" & QuietJson()
End Function

' TASK-P1-164 (A): the live quiet state. The TTL is evaluated on every read, so a mode
' left ON cannot outlive its TTL even if nobody turns it off.
Private Function QuietActive() As Boolean
    If Not m_quietOn Then Exit Function
    If m_quietTtlMs <= 0 Then Exit Function
    If (GetTickCount() - m_quietSince) >= m_quietTtlMs Then
        m_quietOn = False
        Exit Function
    End If
    QuietActive = True
End Function

Private Function QuietJson() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("quietMode") = True
    d("since") = m_quietSince
    d("ttlMs") = m_quietTtlMs
    d("quietExpiresAt") = m_quietSince + m_quietTtlMs
    d("expiresInMs") = m_quietTtlMs - (GetTickCount() - m_quietSince)
    d("suppressedFocusCount") = m_quietSuppressed
    d("advisory") = QUIET_ADVISORY
    d("note") = "while this mode is ON the add-in does not call Window.SetFocus (control_window{action:""activate""}) and gives the foreground back after any menu-driven action (native_breakpoint). foregroundReadable=false means the foreground handle could not be READ in this window station (it is 0 here) - that is NOT evidence that the focus was left alone."
    QuietJson = JsonStringify(d)
End Function

' =====================================================================
' tool dispatch
' =====================================================================

' G-7: is this a source-write permission refusal while the IDE is in break mode?
Private Function IsIdeBusyPermissionError(ByVal msg As String) As Boolean
    On Error Resume Next
    If Len(msg) = 0 Then Exit Function
    Dim st As String
    st = GetIdeState()
    If st <> "break" And st <> "run" Then Exit Function
    If InStr(1, msg, "denied", vbTextCompare) > 0 Then IsIdeBusyPermissionError = True
    If InStr(1, msg, "permission", vbTextCompare) > 0 Then IsIdeBusyPermissionError = True
    If InStr(msg, ChrW$(&H62D2) & ChrW$(&H7EDD)) > 0 Then IsIdeBusyPermissionError = True
    If InStr(msg, ChrW$(&H6743) & ChrW$(&H9650)) > 0 Then IsIdeBusyPermissionError = True
End Function

' ---------------------------------------------------------------------------
' TASK-P1-039 (contract v1.3.2): optional structured error fields on FAILURE receipts.
'   vbErrorCode  <- the VB6 numeric code captured at the Fail handler (add-in side)
'   errorLine    <- the engine position protocol, IDE coords
' Invariants: a field that cannot be obtained is OMITTED (never 0 / never an empty
' string), and a failed read must never disturb the original error path. Every read
' below is wrapped in its own guard for exactly that reason.
' RETRACTED (same card, measured): the raw-global read for `errorLine` lived here and
' read Vba6Base()+&HFC0D188 / +&HFC0D190 through RtlMoveMemory. On the real machine the
' FIRST call that reached it (verify, in break mode) was immediately followed by the IDE
' DYING - the listener disappeared and the VB6 process exited. A raw memory copy is not
' worth an IDE crash, and P2/R64 has not machine-verified those globals yet, so the read
' is REMOVED. `errorLine` is therefore simply ABSENT (contract: a field that cannot be
' obtained is omitted, never faked). `vbErrorCode` below is pure VB data - the number
' captured at the Fail handler - and carries no such risk.
' ---------------------------------------------------------------------------
Private Sub AddErrorFields(ByVal d As Object, ByVal vbErr As Long)
    On Error Resume Next
    If vbErr <> 0 Then d("vbErrorCode") = vbErr
End Sub

' TASK-P1-062: an UNEXPECTED VB runtime error must still come back as a STRUCTURED receipt
' ({ok:false, reason, vbErrorCode}) - the old bare "Error: <text>" gave the caller nothing
' to branch on (no ok, no reason). vbErrorCode keeps the contract v1.3.2 semantics: it is
' present only when a real VB error was captured at the Fail handler.
Private Function UnexpectedErrorJson(ByVal tool As String, ByVal rawErr As String, ByVal vbErr As Long) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "unexpected_error"
    d("tool") = tool
    d("message") = rawErr
    d("advice") = "this is an unexpected VB runtime error in the add-in, not a structured refusal; no claim is made about the component state"
    AddErrorFields d, vbErr
    On Error Resume Next
    d("ideState") = GetIdeState()
    On Error GoTo 0
    UnexpectedErrorJson = JsonStringify(d)
End Function

' ===========================================================================
' TASK-P1-108 (D-1107-01): CLASSIFIED REFUSALS.
'
' D-1107-01 (P1-107, 49 arms / ~34 tools) showed that conditions the add-in has
' ALREADY identified - "File not found: <path>", "Directory not found: <path>",
' "Unknown component type: <t>", "baseline spec is required", "Invalid action: ...",
' "marker begin not found: ...", the VB-native 76/71/32813 codes, and so on - were
' still answered with the BARE envelope ok:false + reason:"unexpected_error" +
' vbErrorCode, which is not branchable without parsing prose.
'
' WHY THEY LEAKED (measured, not guessed): every one of those raises uses
'   Err.Raise vbObjectError + N, "CodeTools", "<classified message>"
' with NO "MCPREASON:" prefix. CallTool's Fail handler only turns a raise into a
' structured refusal when the description starts with "MCPREASON:" (CodeTools.cls
' ~1920); everything else falls through to UnexpectedErrorJson (~1947). So the
' information was never missing - only the channel tag was.
'
' TWO ways to fix: tag ~150 raise sites individually, or classify CENTRALLY here
' from the message family + the captured VB error number. This is the central one:
' one function, one place, and a static completeness scan can verify it against
' every raise site in the product.
'
' SAFETY (contract): these branches run ONLY after the MCPREASON branch and the
' IDE-busy branch, i.e. only on calls that used to answer unexpected_error. A
' successful call never reaches the Fail handler at all, so the success path and
' every existing structured refusal are untouched.
' vbErrorCode is emitted ONLY for a NATIVE error (a real VB/COM code, i.e. one
' outside our own vbObjectError+N band) - TASK-P1-046's rule that our own guards
' must not look like captured VB errors is preserved.
' ===========================================================================
Private Function ClassifiedRefusalJson(ByVal tool As String, ByVal failMsg As String, ByVal vbErr As Long) As String
    On Error Resume Next
    Dim m As String, reason As String, argName As String
    m = failMsg
    reason = ""
    argName = ""
    ' ---- caller-input families (most specific first) -----------------------
    If InStr(1, m, "baseline spec is required", vbTextCompare) > 0 Then
        reason = "required_argument_missing": argName = "baseline"
    ElseIf InStr(1, m, " is required", vbTextCompare) > 0 Then
        reason = "required_argument_missing": argName = ArgNameOf(m)
    ElseIf InStr(1, m, " are required", vbTextCompare) > 0 Then
        reason = "required_argument_missing": argName = ArgNameOf(m)
    ElseIf InStr(1, m, "no fields supplied to set", vbTextCompare) > 0 Then
        reason = "required_argument_missing": argName = "updates"
    ElseIf InStr(1, m, "Project file not found", vbTextCompare) > 0 Then
        reason = "project_file_not_found"
    ElseIf InStr(1, m, "project file is empty or unreadable", vbTextCompare) > 0 Then
        reason = "file_unreadable"
    ElseIf InStr(1, m, "Project not found", vbTextCompare) > 0 Then
        reason = "project_not_found"
    ElseIf InStr(1, m, "No project open", vbTextCompare) > 0 Then
        reason = "project_not_found"
    ElseIf InStr(1, m, "project group not found or empty", vbTextCompare) > 0 Then
        reason = "project_group_not_found"
    ElseIf InStr(1, m, "snapshot.json not found under", vbTextCompare) > 0 Then
        reason = "snapshot_not_found"
    ElseIf InStr(1, m, "Directory not found", vbTextCompare) > 0 Then
        reason = "directory_not_found"
    ElseIf InStr(1, m, "cannot READ the source file", vbTextCompare) > 0 Then
        reason = "file_unreadable"
    ElseIf InStr(1, m, "empty or unreadable", vbTextCompare) > 0 Then
        reason = "file_unreadable"
    ElseIf InStr(1, m, "File not found", vbTextCompare) > 0 Then
        reason = "file_not_found"
    ElseIf InStr(1, m, "procedure not found", vbTextCompare) > 0 Then
        reason = "procedure_not_found"
    ElseIf InStr(1, m, "Invalid operation", vbTextCompare) > 0 Then
        reason = "invalid_action"
    ElseIf InStr(1, m, "selection must be", vbTextCompare) > 0 Then
        reason = "invalid_selection"
    ElseIf InStr(1, m, "looks like an id but its number is missing", vbTextCompare) > 0 Then
        reason = "invalid_control_id"
    ElseIf InStr(1, m, "IDE not connected", vbTextCompare) > 0 Then
        reason = "ide_not_connected"
    ElseIf InStr(1, m, "Unknown component type", vbTextCompare) > 0 Then
        reason = "unknown_component_type"
    ElseIf InStr(1, m, "Unknown project type", vbTextCompare) > 0 Then
        reason = "unknown_project_type"
    ElseIf InStr(1, m, "Unknown attrName", vbTextCompare) > 0 Then
        reason = "unknown_attribute"
    ElseIf InStr(1, m, "Invalid action", vbTextCompare) > 0 Then
        reason = "invalid_action"
    ElseIf InStr(1, m, "Unknown action", vbTextCompare) > 0 Then
        reason = "invalid_action"
    ElseIf InStr(1, m, "Invalid mode", vbTextCompare) > 0 Then
        reason = "invalid_mode"
    ElseIf InStr(1, m, "Invalid level", vbTextCompare) > 0 Then
        reason = "invalid_level"
    ElseIf InStr(1, m, "Reference not found", vbTextCompare) > 0 Then
        reason = "reference_not_found"
    ElseIf InStr(1, m, "Control property not found", vbTextCompare) > 0 Then
        reason = "control_property_not_found"
    ElseIf InStr(1, m, "Property not found", vbTextCompare) > 0 Then
        reason = "property_not_found"
    ElseIf InStr(1, m, "Member not found", vbTextCompare) > 0 Then
        reason = "member_not_found"
    ElseIf InStr(1, m, "Control not found by name", vbTextCompare) > 0 Then
        reason = "control_not_found"
    ElseIf InStr(1, m, "Control index out of range", vbTextCompare) > 0 Then
        reason = "control_index_out_of_range"
    ElseIf InStr(1, m, "index must be >= 1", vbTextCompare) > 0 Then
        reason = "control_index_out_of_range"
    ElseIf InStr(1, m, "Component has no code module", vbTextCompare) > 0 Then
        reason = "component_has_no_code_module"
    ElseIf InStr(1, m, "Component has no designer", vbTextCompare) > 0 Then
        reason = "component_has_no_designer"
    ElseIf InStr(1, m, "Component has no Properties collection", vbTextCompare) > 0 Then
        reason = "component_has_no_properties"
    ElseIf InStr(1, m, "Component has no file name", vbTextCompare) > 0 Then
        reason = "component_has_no_file"
    ElseIf InStr(1, m, "marker begin not found", vbTextCompare) > 0 Then
        reason = "marker_not_found"
    ElseIf InStr(1, m, "anchor not found", vbTextCompare) > 0 Then
        reason = "anchor_not_found"
    ElseIf InStr(1, m, "anchor not unique", vbTextCompare) > 0 Then
        reason = "anchor_not_unique"
    ElseIf InStr(1, m, "import refused", vbTextCompare) > 0 Then
        reason = "import_rebind_failed"
    ElseIf InStr(1, m, "Could not compile", vbTextCompare) > 0 Then
        reason = "compile_failed"
    ElseIf InStr(1, m, "Could not save the project before", vbTextCompare) > 0 Then
        reason = "save_before_build_failed"
    ElseIf InStr(1, m, "Could not create a form component", vbTextCompare) > 0 Then
        reason = "component_add_failed"
    ElseIf InStr(1, m, "VBComponents.Add failed", vbTextCompare) > 0 Then
        reason = "component_add_failed"
    ElseIf InStr(1, m, "References.Remove failed", vbTextCompare) > 0 Then
        reason = "reference_remove_failed"
    ElseIf InStr(1, m, "open_project failed", vbTextCompare) > 0 Then
        reason = "open_project_failed"
    ElseIf InStr(1, m, "build process failed to run", vbTextCompare) > 0 Then
        reason = "build_process_failed"
    ElseIf InStr(1, m, "Add-in not found", vbTextCompare) > 0 Then
        reason = "addin_not_found"
    ElseIf InStr(1, m, "Menu command not found", vbTextCompare) > 0 Then
        reason = "menu_control_not_found"
    ElseIf InStr(1, m, "Menu control not found", vbTextCompare) > 0 Then
        reason = "menu_control_not_found"
    ElseIf InStr(1, m, "Debug menu item not found", vbTextCompare) > 0 Then
        reason = "menu_control_not_found"
    ElseIf InStr(1, m, "no ENABLED menu control", vbTextCompare) > 0 Then
        reason = "menu_control_not_found"
    ElseIf InStr(1, m, "Menu control is disabled", vbTextCompare) > 0 Then
        reason = "menu_control_disabled"
    ElseIf InStr(1, m, "is AMBIGUOUS", vbTextCompare) > 0 Then
        reason = "menu_control_ambiguous"
    ElseIf InStr(1, m, "No active code pane", vbTextCompare) > 0 Then
        reason = "no_active_code_pane"
    ElseIf InStr(1, m, "Project is unsaved", vbTextCompare) > 0 Then
        reason = "project_unsaved"
    ElseIf InStr(1, m, "has no file name", vbTextCompare) > 0 Then
        reason = "project_unsaved"
    ElseIf InStr(1, m, "refusing to save", vbTextCompare) > 0 Then
        reason = "save_refused"
    ElseIf InStr(1, m, "content changed on line", vbTextCompare) > 0 Then
        reason = "file_changed_on_disk"
    ElseIf InStr(1, m, "not a valid identifier", vbTextCompare) > 0 Then
        reason = "invalid_identifier"
    ElseIf InStr(1, m, "old_name and new_name are the same", vbTextCompare) > 0 Then
        reason = "rename_is_noop"
    ElseIf InStr(1, m, "invalid path", vbTextCompare) > 0 Then
        reason = "invalid_path"
    ElseIf InStr(1, m, "Clipboard has nothing pasteable", vbTextCompare) > 0 Then
        reason = "clipboard_empty"
    ElseIf InStr(1, m, "the form has no file on disk yet", vbTextCompare) > 0 Then
        reason = "form_not_saved"
    ElseIf InStr(1, m, "menu item not found in the .frm text", vbTextCompare) > 0 Then
        reason = "menu_item_not_found"
    ElseIf InStr(1, m, "parent menu item not found", vbTextCompare) > 0 Then
        reason = "menu_item_not_found"
    ElseIf InStr(1, m, "already a child of", vbTextCompare) > 0 Then
        reason = "menu_nest_invalid"
    ElseIf InStr(1, m, "cannot nest a menu item inside itself", vbTextCompare) > 0 Then
        reason = "menu_nest_invalid"
    ElseIf InStr(1, m, "could not re-read", vbTextCompare) > 0 Then
        reason = "form_read_failed"
    ElseIf InStr(1, m, "could not find the form's closing End line", vbTextCompare) > 0 Then
        reason = "form_read_failed"
    ElseIf InStr(1, m, "could not read", vbTextCompare) > 0 Then
        reason = "form_read_failed"
    ElseIf InStr(1, m, "Unknown tool", vbTextCompare) > 0 Then
        reason = "unknown_tool"
    ElseIf InStr(1, m, "Could not set", vbTextCompare) > 0 Then
        reason = "property_set_failed"
    ElseIf InStr(1, m, "set_project_options failed", vbTextCompare) > 0 Then
        reason = "set_project_options_failed"
    ElseIf vbErr = 76 Then
        reason = "path_not_found"
    ElseIf vbErr = 71 Then
        reason = "device_not_ready"
    ElseIf vbErr = 53 Then
        reason = "file_not_found"
    ElseIf vbErr = 70 Then
        reason = "permission_denied"
    ElseIf vbErr = 32813 Then
        reason = "com_name_conflict"
    ElseIf vbErr = -2147319779 Then
        reason = "type_not_registered"
    ElseIf vbErr = -2147221164 Then
        reason = "com_class_not_registered"
    End If
    If Len(reason) = 0 Then
        ClassifiedRefusalJson = ""
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("tool") = tool
    d("message") = failMsg
    If Len(argName) > 0 Then d("argument") = argName
    Dim pth As String
    pth = PathOfMessage(m)
    If Len(pth) > 0 Then d("path") = pth
    Dim allowed As Object
    Set allowed = ExpectedListOf(m)
    If Not allowed Is Nothing Then
        If allowed.Count > 0 Then Set d("allowedValues") = allowed
    End If
    If InStr(1, m, "Unknown component type", vbTextCompare) > 0 Then
        Dim ctypes As Object
        Set ctypes = New Collection
        ctypes.Add "module"
        ctypes.Add "class"
        ctypes.Add "form"
        ctypes.Add "mdiform"
        ctypes.Add "designer"
        ctypes.Add "usercontrol"
        ctypes.Add "propPage"
        ctypes.Add "userdoc"
        ctypes.Add "resource"
        Set d("allowedComponentTypes") = ctypes
    End If
    d("suggestedAction") = SuggestedActionFor(reason)
    d("classifiedBy") = "TASK-P1-108 ClassifiedRefusalJson (message family + native VB/COM code allowlist)"
    d("sideEffects") = "none: the call was refused before completing its effect; this is a CLASSIFIED refusal, not an unexpected VB runtime error"
    ' TASK-P1-046 rule preserved: vbErrorCode only for a NATIVE error, never for our own
    ' vbObjectError+N guard raises.
    If Not IsOwnGuardCode(vbErr) Then
        If vbErr <> 0 Then d("vbErrorCode") = vbErr
    End If
    ClassifiedRefusalJson = JsonStringify(d)
End Function

Private Function IsOwnGuardCode(ByVal vbErr As Long) As Boolean
    IsOwnGuardCode = (vbErr >= vbObjectError And vbErr <= vbObjectError + 200)
End Function

' "<arg> is required ..." -> the argument NAME (the text before " is required").
Private Function ArgNameOf(ByVal m As String) As String
    Dim p As Long, s As String, sp As Long
    p = InStr(1, m, " is required", vbTextCompare)
    If p = 0 Then p = InStr(1, m, " are required", vbTextCompare)
    If p = 0 Then
        ArgNameOf = ""
        Exit Function
    End If
    s = Trim$(Left$(m, p - 1))
    If InStr(1, s, " or ", vbTextCompare) > 0 Then
        ArgNameOf = ""
        Exit Function
    End If
    sp = InStr(s, " ")
    If sp > 0 Then s = Left$(s, sp - 1)
    ArgNameOf = s
End Function

' "... (expected a/b/c)" -> a Collection of the advertised alternatives.
Private Function ExpectedListOf(ByVal m As String) As Object
    On Error Resume Next
    Dim i As Long, j As Long, inner As String, parts() As String, k As Long
    Dim c As Object
    Set c = New Collection
    i = InStr(1, m, "(expected ", vbTextCompare)
    If i = 0 Then
        Set ExpectedListOf = Nothing
        Exit Function
    End If
    j = InStr(i, m, ")")
    If j = 0 Then
        Set ExpectedListOf = Nothing
        Exit Function
    End If
    inner = Mid$(m, i + 10, j - i - 10)
    parts = Split(inner, "/")
    For k = 0 To UBound(parts)
        If Len(Trim$(parts(k))) > 0 Then c.Add Trim$(parts(k))
    Next
    Set ExpectedListOf = c
End Function

' the trailing "<label>: <path>" of a not-found message, when it looks like a path.
Private Function PathOfMessage(ByVal m As String) As String
    On Error Resume Next
    Dim p As Long, tail As String
    p = InStrRev(m, ": ")
    If p = 0 Then
        PathOfMessage = ""
        Exit Function
    End If
    tail = Trim$(m)
    tail = Trim$(Mid$(tail, p + 2))
    If Len(tail) = 0 Or Len(tail) > 400 Then
        PathOfMessage = ""
        Exit Function
    End If
    If InStr(tail, ":") = 0 And InStr(tail, Chr$(92)) = 0 And InStr(tail, "/") = 0 Then
        PathOfMessage = ""
        Exit Function
    End If
    PathOfMessage = tail
End Function

Private Function SuggestedActionFor(ByVal reason As String) As String
    Select Case LCase$(reason)
        Case "required_argument_missing"
            SuggestedActionFor = "supply the reported argument (see 'argument' / 'message') and retry"
        Case "file_not_found", "path_not_found"
            SuggestedActionFor = "pass the path of an EXISTING file; a relative path resolves against the project directory, not the process working directory"
        Case "file_unreadable"
            SuggestedActionFor = "check that the file exists, is not locked by another process, and is text in a supported encoding"
        Case "directory_not_found"
            SuggestedActionFor = "pass the path of an EXISTING directory; a relative path resolves against the project directory"
        Case "project_file_not_found", "project_not_found"
            SuggestedActionFor = "pass the .vbp of an existing project, or open the project in the IDE first"
        Case "unknown_component_type", "unknown_project_type", "unknown_attribute"
            SuggestedActionFor = "use one of allowedValues (or allowedComponentTypes) and retry"
        Case "invalid_action", "invalid_mode", "invalid_level"
            SuggestedActionFor = "use one of allowedValues and retry"
        Case "reference_not_found"
            SuggestedActionFor = "call check_references and use one of the listed reference names"
        Case "control_not_found", "property_not_found", "control_property_not_found", "member_not_found"
            SuggestedActionFor = "call the matching inspector (form_designer action=info / list_members / get_component_properties) and use a name it lists"
        Case "component_has_no_code_module"
            SuggestedActionFor = "this component kind has no code pane (a designer-only component); pick a module/class/form"
        Case "component_has_no_designer"
            SuggestedActionFor = "this component kind has no designer; use form_designer only on a form/mdiform"
        Case "procedure_not_found"
            SuggestedActionFor = "call list_members (or get_component_list) and use a procedure name it lists"
        Case "invalid_selection"
            SuggestedActionFor = "pass selection as startLine,startCol,endLine,endCol (four comma-separated numbers)"
        Case "invalid_control_id"
            SuggestedActionFor = "address the control by index, or write the id as #746 / id:746 so the number is present"
        Case "ide_not_connected"
            SuggestedActionFor = "the add-in has no live VBIDE connection; start/attach the IDE and retry"
        Case "anchor_not_found", "anchor_not_unique"
            SuggestedActionFor = "make the anchor match exactly one pane line (see get_component_code coords='ide')"
        Case "marker_not_found"
            SuggestedActionFor = "the injected marker with that id is not present; list_breakpoints / the matching list tool shows what exists"
        Case "project_unsaved"
            SuggestedActionFor = "call save_project (with a path when the project has no file name) and retry"
        Case "device_not_ready", "path_not_found_native"
            SuggestedActionFor = "the drive or path is not available to this process; check the drive letter and retry"
        Case "com_name_conflict"
            SuggestedActionFor = "a component with that name already exists in the project; rename it before adding the reference"
        Case "type_not_registered", "com_class_not_registered"
            SuggestedActionFor = "the type library / COM class is not registered for the 32-bit IDE; register it with %SystemRoot% SysWOW64 regsvr32"
        Case "compile_failed", "save_before_build_failed", "build_process_failed"
            SuggestedActionFor = "read the compile log reported by build/make_project; fix the reported compile error and retry"
        Case "menu_control_not_found", "menu_control_disabled", "menu_control_ambiguous"
            SuggestedActionFor = "call command_bars action=list and address the control by bar= / caption / id= as it is listed"
        Case "no_active_code_pane"
            SuggestedActionFor = "activate a component's code pane first (activate_component window='code')"
        Case Else
            SuggestedActionFor = "fix the reported condition (see 'reason' and 'message') and retry"
    End Select
End Function

Private Function IdeBusyBreakJson(ByVal tool As String, ByVal rawErr As String, Optional ByVal vbErr As Long = 0) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim st As String
    st = GetIdeState()
    d("ok") = False
    If st = "break" Then
        d("reason") = "ide_busy_in_break_mode"
        d("advice") = "the debugger is stopped (break mode) and the IDE refuses source writes; call control_execution with mode=end to leave break mode, then retry"
    Else
        d("reason") = "ide_busy_in_run_mode"
        d("advice") = "the program is running and the IDE refuses source writes; call control_execution with mode=end (or break first) and retry in design mode"
    End If
    d("ideState") = st
    d("tool") = tool
    d("error") = rawErr
    AddErrorFields d, vbErr
    IdeBusyBreakJson = JsonStringify(d)
End Function

' ===========================================================================
' TASK-P1-066: re-entrancy gate around tool dispatch.
'
' DoEvents hands the IDE message queue back, which means a NEW HTTP request can
' be dispatched THROUGH THIS SAME THREAD while a tool call is mid-flight. That
' is exactly what makes the IDE responsive, and exactly what makes an
' unguarded write path dangerous: two state-changing calls would interleave
' across the pump point (call A reads the pane, call B rewrites it, call A
' writes back its stale prediction).
'
' Policy (chosen over queueing: VB6 is single threaded and this add-in has no
' request queue, so a refusal is predictable and testable):
'   * while a STATE-CHANGING tool is in flight, every re-entrant call from a
'     pump is answered at once with {ok:false, reason:"busy", busyTool,...} -
'     never executed, so two mutations can never interleave;
'   * while only READ-ONLY tools are in flight, a re-entrant call IS served
'     (that is the responsiveness goal). A nested state-changing call is then
'     allowed, and the outer read-only loop re-reads the model on every
'     iteration, so it cannot act on a stale pane;
'   * nesting is capped (CALL_DEPTH_MAX) so a pathological chain of pumped
'     requests cannot grow the stack without bound.
'
' The mutating classification is a conservative ALLOWLIST of read-only tools:
' anything not listed - including every extension tool - counts as
' state-changing. A read-only tool that is missing from the list merely gets
' refused during a mutation; the failure mode is never "two writes interleaved".
' ===========================================================================

Private Function IsMutatingTool(ByVal name As String) As Boolean
    Select Case LCase$(name)
        Case "ide_state", "current_position", "list_breakpoints", "ui_text", "find_gui", _
             "get_project_info", "get_component_list", "get_component_code", _
             "find_in_component", "get_codepane", "get_selection", "list_files", _
             "read_file", "search_in_files", "list_component_parts", "list_addins", _
             "validate_code", "read_source", "check_component", "preview_edit", "verify", _
             "audit_tail", "debug_commands", "execution_state", "vba6_probe_plan", _
             "export_public_api", "check_public_api", "purity_check", "compare_fixture", _
             "list_lessons", "export_lessons", "ext_list", "project_changes", _
             "get_callstack", "check_references", "workspace_status", "verify_scope", _
             "todo_status", "mcp_info", "list_members", "get_component_properties", "git", _
             "dependency_graph", "control_execution", "read_active_dialogs", _
             "ide_quiet_mode"
            IsMutatingTool = False
        Case Else
            IsMutatingTool = True
    End Select
End Function

Private Function TextResult(ByVal body As String, ByVal isErr As Boolean) As Object
    Dim res As Object, content As Object, item As Object
    Set res = CreateObject("Scripting.Dictionary")
    Set content = New Collection
    Set item = CreateObject("Scripting.Dictionary")
    item("type") = "text"
    item("text") = StampMcp(body)
    content.Add item
    Set res("content") = content
    res("isError") = isErr
    Set TextResult = res
End Function

Private Function BusyRejectResult(ByVal name As String, ByVal busyTool As String, ByVal busyAt As Long, _
                                  Optional ByVal inflightTool As String = "", Optional ByVal busyKind As String = "write", Optional ByVal readonlyOk As Boolean = False) As Object
    ' TASK-P1-125 (MCP-01 / MCP-08 / MCP-09): busyTool + busyMs alone do not let a
    ' caller decide "wait" versus "restart the IDE". The fields below make the
    ' in-flight call machine readable, state that only the WRITE class is blocked
    ' (MCP-09 recorded that read-only tools keep answering - that strength is
    ' preserved and is now documented in the receipt itself), and give an
    ' executable way out instead of leaving the caller to kill the IDE.
    Dim d As Object
    Dim ms As Long
    ms = GetTickCount() - busyAt
    If Len(inflightTool) = 0 Then inflightTool = busyTool
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "busy"
    d("tool") = name
    d("busyTool") = busyTool
    d("busyMs") = ms
    d("inflight_tool") = inflightTool
    d("inflight_since_ms") = ms
    d("inflight_serviceable") = False
    d("inflight_kind") = busyKind
    Dim inFlight As Object
    Set inFlight = CreateObject("Scripting.Dictionary")
    inFlight("active") = True
    inFlight("tool") = inflightTool
    inFlight("sinceMs") = ms
    inFlight("serviceable") = False
    inFlight("kind") = busyKind
    Set d("inflight") = inFlight
    ' TASK-P1-125 (S point 4): the A-side measurement decides this, not a slogan.
    ' While a call is already in flight through the pump point, an incoming READ-ONLY
    ' call is refused too (measured: read_active_dialogs returned this same receipt),
    ' so readonly_serviceable must be FALSE here. It is TRUE only in the case where a
    ' read-only OUTER call is in flight and a WRITE was refused - there a read-only
    ' incoming call still passes the gate.
    If readonlyOk Then
        d("readonly_serviceable") = True
        d("readonly_serviceable_note") = "in THIS case a read-only call is still served: only a WRITE is refused while a read-only outer call is in flight through the pump point"
    Else
        d("readonly_serviceable") = False
        d("readonly_serviceable_note") = "in THIS case a READ-ONLY call is refused as well, because it arrived at depth>0 through the same message-pump point while a call was already in flight (measured live: read_active_dialogs got exactly this receipt). Retry from a fresh call once the in-flight one returns."
    End If
    d("readonly_tools_note") = "read-only TOOLS are not gated by m_mutBusy when they arrive at depth 0, so they keep answering while a WRITE is in flight (MCP-09, measured on the A-side: ide_state answered in 16 ms while a write was busy). This does NOT make writes available: a WRITE is refused while ANY call is in flight through a message-pump point, INCLUDING a read-only outer call (A-side measured: busyTool='validate_code (outer call, in flight since 141 ms ago)'). inflight_tool / inflight_kind / inflight_since_ms say who is in flight and whether it is a read or a write."
    d("write_refused_because") = "a call is already in flight through this add-in's message pump: serving a second call there would interleave two mutations (or two readers), so it is refused instead of run"
    d("inflight_cancel") = "not available: this add-in cannot abort a call that is already running inside the host. See recovery_hint."
    d("recovery_hint") = "1) WAIT - a normal call returns within its own timeout, after which this refusal clears itself (in flight " & CStr(ms \ 1000) & "s so far). 2) INSPECT WITHOUT WRITING - call read_active_dialogs (read-only) to see whether the host is blocked on a modal box (ownerDisabled=true). 3) IF inflight_since_ms KEEPS GROWING past that call's own timeout, the HOST is inside a modal loop (a clicked button opened VB6's own modal): the add-in cannot interrupt it, so restart the IDE - end the VB6 process, then start it again on the same .vbp; the add-in re-attaches on load and nothing in the project directory is lost."
    d("advice") = "another state-changing tool call is in flight (" & busyTool & "). This add-in is single threaded; serving a second call from a message-pump point would interleave two mutations, so it is refused instead of run. Retry after it returns."
    d("pump") = PumpStatsJson()
    Set BusyRejectResult = TextResult(SealedReceipt(name, JsonStringify(d)), True)
End Function

' ===================== TASK-P1-126 (MCP-08 / MCP-09) =====================
' inflight{} - the machine-readable in-flight block. It is spliced into EVERY
' read-only receipt (SealedReceipt) and every busy refusal (BusyRejectResult).
' Semantics, fixed by measurement rather than by slogan:
'   active     : a call OTHER than the one producing this receipt is in flight
'   tool       : that call's tool name
'   sinceMs    : how long it has been in flight
'   serviceable: a READ-ONLY call can still be served right now
'   kind       : write | read | nested | none
'   self*      : who produced THIS receipt
' =======================================================================
Private Function InflightJson() As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim act As Boolean, who As String, since As Long, kind As String, svc As Boolean
    act = False
    who = ""
    since = 0
    kind = "none"
    svc = True
    If m_mutBusy And m_callDepth > 1 Then
        act = True
        who = m_mutBusyTool
        since = GetTickCount() - m_mutBusyAt
        kind = "write"
    ElseIf m_callDepth > 1 Then
        act = True
        who = m_callToolName
        since = GetTickCount() - m_callStartAt
        kind = "read"
    End If
    d("active") = act
    d("tool") = who
    d("sinceMs") = since
    d("serviceable") = svc
    d("kind") = kind
    d("depth") = m_callDepth
    d("self_tool") = m_callToolName
    d("self_kind") = IIf(m_callIsMut, "write", "read")
    d("cancelRequested") = m_cancelRequested
    d("readonlyWhileWrite") = (m_mutBusy And Not m_callIsMut)
    d("stampedCount") = m_inflightStamped
    ' TASK-P1-168 (C): once the gate has EVER been force-released, every read-only
    ' receipt says so - the caller must be able to see that the single-flight guarantee
    ' was traded away, and for which call. While the feature is unused these keys are
    ' absent, so receipts stay byte-for-byte the pre-v1.9.0 ones.
    If m_gateReleaseCount > 0 Then
        d("gateReleased") = True
        d("gateReleaseCount") = m_gateReleaseCount
        d("gateReleasedBy") = m_gateReleasedBy
        d("gateReleasedTool") = m_gateReleasedTool
        d("gateReleasedHeldMs") = m_gateReleasedHeldMs
        d("gateReleasedAt") = m_gateReleasedAt
        d("gateReleaseReason") = m_gateReleasedReason
        d("gateReleaseNote") = "the gate was force-released for a call that never came back; that call is STILL RUNNING inside the IDE, so a write issued now can interleave with it. Dismiss the host modal (close_active_dialogs) to let it finish."
    End If
    d("note") = "read-only calls ARE served while a WRITE is in flight (MCP-09); a WRITE is refused while ANY call is in flight through the message pump (MCP-08). A read served while a write is in flight can observe a mid-mutation snapshot, which is why active=true and kind=write are reported with it."
    InflightJson = JsonStringify(d)
    Err.Clear
End Function

' TASK-P1-126 (MCP-08): build the cancel{} block of an ide_state receipt. The KEY SET
' is fixed and identical on every path (discipline 98); every un-applied path also
' carries a reason.
Private Function CancelRequestBody(ByVal byTool As String, ByVal targetTool As String, ByVal requested As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("requested") = requested
    d("applied") = False
    d("targetTool") = ""
    d("targetKind") = ""
    d("sinceMs") = 0
    d("depth") = m_callDepth
    d("reason") = ""
    d("honorMinMs") = CANCEL_HONOR_MIN_MS
    d("note") = ""
    If Not requested Then
        d("reason") = "not_requested"
        d("note") = "pass cancel_inflight=true to ask the in-flight call to stop"
        Set CancelRequestBody = d
        Exit Function
    End If
    If m_callDepth <= 1 Then
        d("reason") = "nothing_in_flight"
        d("note") = "no OTHER call is in flight (this read-only call is the only one), so there is nothing to cancel"
        Set CancelRequestBody = d
        Exit Function
    End If
    Dim tt As String, tk As String
    tt = targetTool
    tk = ""
    If Len(tt) = 0 Then
        If m_mutBusy Then
            tt = m_mutBusyTool
            tk = "write"
        Else
            tt = m_callToolName
            tk = "read"
        End If
    End If
    If m_mutBusy Then
        d("sinceMs") = GetTickCount() - m_mutBusyAt
    Else
        d("sinceMs") = GetTickCount() - m_callStartAt
    End If
    m_cancelRequested = True
    m_cancelAt = GetTickCount()
    m_cancelByTool = byTool
    m_cancelTargetTool = tt
    m_cancelTargetKind = tk
    d("applied") = True
    d("targetTool") = tt
    d("targetKind") = tk
    d("reason") = "flag_set"
    d("note") = "the flag is honoured at the in-flight call's next pump point (at least " & CStr(CANCEL_HONOR_MIN_MS) & " ms into the call). A call already blocked inside a HOST modal loop never reaches one of our pump points and cannot see it - use the hub tool ide_restart for that case."
    Set CancelRequestBody = d
End Function

' TASK-P1-126 (MCP-08): called from the pump loops of the LONG calls. Only the
' OUTERMOST call (depth 1) honours a cancel, so a nested reader can never abort
' itself. Deliberately NOT wrapped in On Error Resume Next: the raise below IS the
' abort.
Private Sub MaybeHonorCancel()
    If Not m_cancelRequested Then Exit Sub
    If m_callDepth <> 1 Then Exit Sub
    If Len(m_cancelTargetTool) = 0 Then Exit Sub
    If (GetTickCount() - m_cancelAt) < CANCEL_HONOR_MIN_MS Then Exit Sub
    Err.Raise vbObjectError + 91, "CodeTools", "MCPREASON:cancelled_by_request|" & m_cancelTargetTool & "|" & m_cancelTargetKind
End Sub

' TASK-P1-168 (C / defect #10): FORCE-RELEASE the single-flight gate.
'   targetDepth = the depth to KEEP. A caller that is itself executing INSIDE the gate
'   (ide_state) passes 1 so its own frame stays counted; a caller that runs BEFORE its own
'   depth is incremented (CallTool entry, i.e. the armed release) passes 0.
'   The returned Dictionary has a FIXED key set on every path, so a caller can branch on
'   reason/released without probing for keys.
Private Function ForceReleaseGate(ByVal byTool As String, ByVal why As String, ByVal minHeldMs As Long, ByVal targetDepth As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("attempted") = True
    d("released") = False
    d("releasedTool") = ""
    d("releasedKind") = ""
    d("heldMs") = 0
    d("minHeldMs") = minHeldMs
    d("retryAfterMs") = 0
    d("depthBefore") = m_callDepth
    d("depthAfter") = m_callDepth
    d("targetDepth") = targetDepth
    d("mutBusyBefore") = m_mutBusy
    d("abandonedSeq") = 0
    d("reason") = ""
    d("gateReleaseCount") = m_gateReleaseCount
    d("statementInterrupted") = False
    d("timeoutMsEnforcedBy") = "gate_age_threshold"
    d("note") = ""
    d("recoveryAction") = ""
    d("recoveryHint") = ""
    If m_callDepth <= targetDepth Then
        d("reason") = "nothing_in_flight"
        d("note") = "no OTHER call is in flight (this read-only call is the only one), so there is nothing to release"
        Set ForceReleaseGate = d
        Exit Function
    End If
    Dim tk As String, at As Long
    tk = ""
    at = 0
    If m_mutBusy Then
        tk = m_mutBusyTool
        at = m_mutBusyAt
    Else
        tk = m_outerTool
        at = m_outerAt
    End If
    Dim held As Long
    held = GetTickCount() - at
    d("releasedTool") = tk
    d("releasedKind") = IIf(m_mutBusy, "write", "read")
    d("heldMs") = held
    If minHeldMs < 0 Then minHeldMs = 0
    d("minHeldMs") = minHeldMs
    If held < minHeldMs Then
        d("reason") = "age_below_threshold"
        d("retryAfterMs") = minHeldMs - held
        d("note") = "the in-flight call is younger than minHeldMs, so NOTHING was touched. This threshold is what keeps a HEALTHY long call (a normal build, a slow write) from losing the single-flight guarantee."
        d("recoveryAction") = "call ide_state{release_inflight:true, release_min_held_ms:<ms>} again after retryAfterMs"
        d("recoveryHint") = "wait retryAfterMs and repeat the call"
        Set ForceReleaseGate = d
        Exit Function
    End If
    ' ---- the release ----
    m_gateReleaseCount = m_gateReleaseCount + 1
    m_gateReleasedBy = byTool
    m_gateReleasedTool = tk
    m_gateReleasedHeldMs = held
    m_gateReleasedAt = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    m_gateReleasedReason = why
    m_abandonSeq = m_outerSeq
    m_callDepth = targetDepth
    m_mutBusy = False
    m_mutBusyTool = ""
    m_mutBusyAt = 0
    m_armMs = 0
    m_armTool = ""
    m_armAt = 0
    d("released") = True
    d("abandonedSeq") = m_abandonSeq
    d("depthAfter") = m_callDepth
    d("gateReleaseCount") = m_gateReleaseCount
    d("reason") = why
    d("note") = "the blocked call is STILL RUNNING inside the IDE; the gate was force-released and the work it was doing was NOT undone. This add-in cannot preempt EbExecuteLine, so no statement was interrupted."
    d("recoveryAction") = "close_active_dialogs"
    d("recoveryHint") = "the blocked call usually waits on the IDE's OWN modal run-time-error box (read_active_dialogs shows it with ownerDisabled=true). Dismiss that box (close_active_dialogs) so the blocked call can return - its unwind is swallowed exactly once, so the depth can never be decremented twice. If no modal is present the statement is simply still running: use the hub tool ide_restart, or wait."
    Set ForceReleaseGate = d
End Function

Public Function CallTool(ByVal name As String, ByVal args As Object) As Object
    ' TASK-P1-066: the gate. m_callDepth > 0 means this call arrived from a
    ' message-pump point inside another tool call (a normal call always starts
    ' at depth 0).
    ' TASK-P1-168 (C): every call carries a token. The token is what makes the forced
    ' release safe: only the frame whose token was abandoned skips its unwind.
    m_callSeq = m_callSeq + 1
    Dim mySeq As Long
    mySeq = m_callSeq
    If m_callDepth > 0 Then
        ' TASK-P1-168 (B/C): an ARMED immediate_* call that outlived its own timeout_ms is
        ' released HERE, before the gate is consulted, so THIS call is the one that finds the
        ' session usable again instead of being refused with reason=busy.
        MaybeAutoRelease
    End If
    If m_callDepth > 0 Then
        ' (1) a state-changing call is ALREADY in flight.
        ' TASK-P1-126 (MCP-09 kept / MCP-08 answered): a WRITE is STILL refused -
        ' two mutations must never interleave - but a READ-ONLY call is now SERVED.
        ' That is the strength the user recorded (ide_state answered in 16 ms while a
        ' write was busy). P1-125 measured the opposite for the make_project case (the
        ' read was refused too); nothing is hidden by this change: the served read
        ' carries inflight{active:true,kind:"write"}, so the caller can see it may be
        ' reading a mid-mutation snapshot.
        If m_mutBusy Then
            If IsMutatingTool(name) Then
                Set CallTool = BusyRejectResult(name, m_mutBusyTool, m_mutBusyAt, m_mutBusyTool, "write", False)
                Exit Function
            End If
        End If
        ' (2) THIS call changes state and would start inside another call's pump.
        '     It is refused even when the outer call is read-only: a write must
        '     never run re-entrantly, and the outer read-only loop would otherwise
        '     keep reporting a snapshot the write has already invalidated.
        If IsMutatingTool(name) Then
            Set CallTool = BusyRejectResult(name, m_callToolName & " (outer call, in flight since " & CStr(GetTickCount() - m_callStartAt) & " ms ago)", m_callStartAt, m_callToolName, "read", True)
            Exit Function
        End If
        ' (3) bound the nesting depth of pumped read-only requests.
        If m_callDepth >= CALL_DEPTH_MAX Then
            Set CallTool = BusyRejectResult(name, "nested depth " & CStr(m_callDepth), GetTickCount(), m_callToolName, "nested", False)
            Exit Function
        End If
    End If
    Dim wasMut As Boolean
    wasMut = IsMutatingTool(name)
    Dim prevSelfMut As Boolean
    prevSelfMut = m_callIsMut
    m_callDepth = m_callDepth + 1
    m_callToolName = name
    m_callStartAt = GetTickCount()
    m_callIsMut = wasMut
    ' TASK-P1-085: one reset point for the whole call, at the only funnel every tool goes
    ' through, so no receipt can inherit the previous call's project-directory writes.
    ' TASK-P1-126: only the OUTERMOST call may reset them. A nested read-only call
    ' (served at the pump point while a write is in flight, see the gate above) used to
    ' wipe the OUTER call's sideEffects/undo manifest half-way through the write.
    If m_callDepth = 1 Then
        ' TASK-P1-168 (C): the OUTERMOST frame's identity, captured BEFORE a nested read can
        ' overwrite m_callToolName / m_callStartAt (it sets both for ITSELF), so a forced
        ' release can never name the wrong tool.
        m_outerSeq = mySeq
        m_outerTool = name
        m_outerAt = GetTickCount()
        Set m_dirWrites = New Collection
        Set m_paneDiskMaps = CreateObject("Scripting.Dictionary")
        SeReset
    End If
    If wasMut Then
        m_mutBusy = True
        m_mutBusyTool = name
        m_mutBusyAt = GetTickCount()
    End If
    PumpReset
    On Error GoTo WrapperFail
    Set CallTool = CallToolInner(name, args)
    On Error GoTo 0
    CallUnwind name, mySeq, wasMut, prevSelfMut
    Exit Function
WrapperFail:
    Dim wErr As Long, wMsg As String
    wErr = Err.Number
    wMsg = Err.Description
    Err.Clear
    CallUnwind name, mySeq, wasMut, prevSelfMut
    Set CallTool = TextResult(SealedReceipt(name, UnexpectedErrorJson(name, wMsg, wErr)), True)
End Function

' TASK-P1-168 (C): the ONE unwind point. A frame whose token was abandoned by a forced
' release must NOT unwind again - the release already dropped it from m_callDepth, and
' decrementing a second time would leave a phantom outer frame that makes later calls be
' refused with reason=busy intermittently (the exact class of bug the lead asked to rule out).
Private Sub CallUnwind(ByVal name As String, ByVal mySeq As Long, ByVal wasMut As Boolean, ByVal prevSelfMut As Boolean)
    If mySeq <> 0 And mySeq = m_abandonSeq Then
        ' Swallow this unwind exactly once. The CURRENT in-flight accounting (a newer call)
        ' is left untouched.
        m_abandonSeq = 0
        Exit Sub
    End If
    m_callDepth = m_callDepth - 1
    ' Only the call that OWNS the flag may clear it: an abandoned frame can return while a
    ' NEW write is in flight, and clearing it there would let a third call interleave.
    If wasMut Then
        If m_mutBusyTool = name Then m_mutBusy = False
    End If
    m_callIsMut = prevSelfMut
    If m_callDepth = 0 Then m_cancelRequested = False
End Sub

' TASK-P1-168 (B): the armed, SELF-HEALING release. immediate_eval / immediate_run arm it
' immediately before handing their line to EbExecuteLine and disarm it on a normal return.
' If the engine never comes back, the arm survives and the next tool call releases the gate
' on its own, so one bad statement cannot leave the session permanently busy.
Private Sub MaybeAutoRelease()
    If m_armMs <= 0 Then Exit Sub
    If Not m_mutBusy Then Exit Sub
    If Len(m_armTool) = 0 Then Exit Sub
    If m_mutBusyTool <> m_armTool Then Exit Sub
    If (GetTickCount() - m_mutBusyAt) < m_armMs Then Exit Sub
    ' byTool says WHO/WHAT released it: "armed_release" is the automatic arm, NOT an
    ' ide_state call that carried release_inflight=true. Labelling it "ide_state" made the
    ' receipt claim a request nobody made (caught on the live run of this very card).
    Call ForceReleaseGate("armed_release", "auto_released_armed_timeout", m_armMs, 0)
End Sub

Private Function CallToolInner(ByVal name As String, ByVal args As Object) As Object
    Dim res As Object
    Set res = CreateObject("Scripting.Dictionary")
    Dim content As Object
    Set content = New Collection
    Dim item As Object
    Set item = CreateObject("Scripting.Dictionary")
    item("type") = "text"
    Dim outText As String
    On Error GoTo Fail
    Select Case name
        Case "get_project_info"
            outText = ToolGetProjectInfo()
        Case "set_project_options"
            outText = ToolSetProjectOptions( _
                JsonGetStr(args, "description"), JsonGetStr(args, "help_file"), _
                JsonGetLong(args, "help_context_id"), JsonGetStr(args, "startup_object"), _
                JsonGetStr(args, "start_project"), JsonGetStr(args, "build_file_name"), _
                JsonGetStr(args, "active_project"), JsonGetBool(args, "read_only_mode", False), _
                JsonGetBool(args, "read_only_mode_set", False))
        Case "open_project"
            outText = ToolOpenProject(JsonGetStr(args, "path"), JsonGetBool(args, "exclusive", False))
        Case "add_project"
            outText = ToolAddProject(JsonGetStr(args, "type"), JsonGetStr(args, "template_path"))
        Case "remove_project"
            outText = ToolRemoveProject(JsonGetStr(args, "name"))
        Case "get_component_list"
            outText = ToolGetComponentList()
        Case "get_component_code"
            outText = ToolGetComponentCode(JsonGetStr(args, "name"), JsonGetBool(args, "include_line_numbers"), JsonGetStr(args, "coords"), JsonGetStr(args, "proc"), JsonGetStr(args, "symbol"))
        Case "replace_code"
            outText = ToolReplaceCode(JsonGetStr(args, "name"), JsonGetStr(args, "operation"), JsonGetLong(args, "start_line"), JsonGetLong(args, "end_line"), JsonGetStr(args, "code"), JsonGetStr(args, "expected_sha256"), JsonGetStr(args, "validate"), SnapshotOptIn(args))
        Case "add_component"
            outText = ToolAddComponent(JsonGetStr(args, "type"), JsonGetStr(args, "name"), JsonGetStr(args, "code"))
        Case "remove_component"
            outText = ToolRemoveComponent(JsonGetStr(args, "name"))
        Case "import_component"
            outText = ToolImportComponent(JsonGetStr(args, "path"), JsonGetStr(args, "mode", "file"))
        Case "export_component"
            outText = ToolExportComponent(JsonGetStr(args, "name"), JsonGetStr(args, "path"))
        Case "rename_component"
            outText = ToolRenameComponent(JsonGetStr(args, "name"), JsonGetStr(args, "new_name"))
        Case "reload_component"
            outText = ToolReloadComponent(JsonGetStr(args, "name"))
        Case "add_code"
            outText = ToolAddCode(JsonGetStr(args, "component"), JsonGetStr(args, "code"), JsonGetStr(args, "file"), JsonGetStr(args, "validate"))
        Case "create_event_proc"
            outText = ToolCreateEventProc(JsonGetStr(args, "component"), JsonGetStr(args, "object"), JsonGetStr(args, "event"))
        Case "find_in_component"
            outText = ToolFindInComponent(JsonGetStr(args, "component"), JsonGetStr(args, "target"), JsonGetLong(args, "start_line"), JsonGetLong(args, "start_col"), JsonGetLong(args, "end_line"), JsonGetLong(args, "end_col"), JsonGetBool(args, "whole_word"), JsonGetBool(args, "match_case"), JsonGetBool(args, "pattern_search"))
        Case "get_codepane"
            outText = ToolGetCodePane(JsonGetStr(args, "component"))
        Case "set_codepane"
            outText = ToolSetCodePane(JsonGetStr(args, "component"), JsonGetStr(args, "selection"), JsonGetLong(args, "top_line"), JsonGetStr(args, "view"), JsonGetBool(args, "activate"))
        Case "get_selection"
            outText = ToolGetSelection()
        Case "activate_component"
            outText = ToolActivateComponent(JsonGetStr(args, "name"), JsonGetStr(args, "window", "activate"))
        Case "list_files"
            outText = ToolListFiles(JsonGetStr(args, "path"), JsonGetStr(args, "pattern"), JsonGetBool(args, "recursive", True), JsonGetLong(args, "max_entries", 500))
        Case "read_file"
            outText = ToolReadFile(JsonGetStr(args, "path"), JsonGetLong(args, "max_lines", 0))
        Case "write_file"
            outText = ToolWriteFile(JsonGetStr(args, "path"), JsonGetStr(args, "content"), JsonGetStr(args, "encoding", "auto"), JsonGetBool(args, "create_dirs", True), JsonGetBool(args, "backup", True), JsonGetStr(args, "validate"))
        Case "search_in_files"
            outText = ToolSearchInFiles(JsonGetStr(args, "query"), JsonGetStr(args, "path"), JsonGetStr(args, "pattern"), JsonGetBool(args, "is_regex"), JsonGetBool(args, "case_sensitive"), JsonGetLong(args, "max_results", 200))
        Case "save_project"
            outText = ToolSaveProject(JsonGetStr(args, "path"))
        Case "make_project"
            outText = ToolMakeProject(JsonGetStr(args, "output_path"), JsonGetStr(args, "route"), JsonGetStr(args, "out_path"), JsonGetLong(args, "timeout_ms"), JsonGetBool(args, "allow_host_modal", False), HasArg(args, "laa"), JsonGetBool(args, "laa", True), JsonGetBool(args, "diagnose", False))
        Case "control_execution"
            outText = ToolControlExecution(JsonGetStr(args, "mode"))
        Case "test_host"
            outText = ToolTestHost(JsonGetStr(args, "action"), JsonGetStr(args, "vbp_path"), JsonGetStr(args, "module_name"), _
                                   JsonGetStr(args, "module_under_test"), JsonGetStr(args, "assertions_path"), _
                                   JsonGetStr(args, "runner_path"), JsonGetStr(args, "exe_name"), JsonGetStr(args, "result_path"))
        Case "project_meta"
            outText = ToolProjectMeta(args)
        Case "app_manifest"
            outText = ToolAppManifest(args)
        Case "check_references"
            outText = ToolCheckReferences(args)
        Case "vbg_group"
            outText = ToolVbgGroup(args)
        Case "project_snapshot"
            outText = ToolProjectSnapshot(args)
        Case "package_app"
            outText = ToolPackageApp(args)
        Case "git"
            outText = ToolGit(args)
        Case "format_code"
            outText = ToolFormatCode(args)
        Case "rename_symbol"
            outText = ToolRenameSymbol(args)
        Case "add_reference"
            outText = ToolAddReference(JsonGetStr(args, "mode", "guid"), JsonGetStr(args, "guid"), JsonGetStr(args, "path"), JsonGetLong(args, "major", 0), JsonGetLong(args, "minor", 0))
        Case "remove_reference"
            outText = ToolRemoveReference(JsonGetStr(args, "name"))
        Case "get_component_properties"
            outText = ToolGetComponentProperties(JsonGetStr(args, "name"))
        Case "set_component_property"
            outText = ToolSetComponentProperty(JsonGetStr(args, "name"), JsonGetStr(args, "property"), JsonGetStr(args, "value"), JsonGetStr(args, "indices"))
        Case "list_members"
            outText = ToolListMembers(JsonGetStr(args, "component"))
        Case "set_member_attribute"
            outText = ToolSetMemberAttribute(JsonGetStr(args, "component"), JsonGetStr(args, "member"), JsonGetStr(args, "attribute"), JsonGetStr(args, "value"))
        Case "list_windows"
            outText = ToolListWindows()
        Case "control_window"
            outText = ToolControlWindow(JsonGetStr(args, "window"), JsonGetStr(args, "action"), JsonGetStr(args, "caption"), JsonGetLong(args, "left"), JsonGetLong(args, "top"), JsonGetLong(args, "width"), JsonGetLong(args, "height"), HasArg(args, "caption"))
        Case "form_designer"
            outText = ToolFormDesigner(JsonGetStr(args, "form"), JsonGetStr(args, "action"), JsonGetStr(args, "control"), JsonGetLong(args, "index"), JsonGetStr(args, "parent"))
        Case "add_control"
            outText = ToolAddControl(JsonGetStr(args, "form"), JsonGetStr(args, "prog_id"), JsonGetStr(args, "name"))
        Case "list_component_parts"
            outText = ToolListComponentParts()
        Case "remove_control"
            outText = ToolRemoveControl(JsonGetStr(args, "form"), JsonGetLong(args, "index"), JsonGetStr(args, "control"))
        Case "set_control_property"
            outText = ToolSetControlProperty(JsonGetStr(args, "form"), JsonGetLong(args, "index"), JsonGetStr(args, "control"), JsonGetStr(args, "property"), JsonGetStr(args, "value"))
        Case "native_breakpoint"
            outText = ToolNativeBreakpoint(args)
        Case "property_bag"
            ' TASK-P1-112: the schema's own \`value\` description promises "Value to write (omit to
            ' read)", yet the write trigger was the UNDECLARED has_value boolean - so a caller who
            ' followed the documentation passed value and got a silent READ back (the exact defect).
            ' The PRESENCE of \`value\` now selects a WRITE, exactly as documented; has_value stays
            ' honoured as an explicit override and is now DECLARED in the schema (registration below).
            outText = ToolPropertyBag(JsonGetStr(args, "target"), JsonGetStr(args, "name"), JsonGetStr(args, "section"), JsonGetStr(args, "key"), JsonGetStr(args, "value"), WantsPropertyWrite(args))
        Case "list_addins"
            outText = ToolListAddins()
        Case "set_addin_state"
            outText = ToolSetAddinState(JsonGetStr(args, "prog_id"), JsonGetBool(args, "connect", True))
        Case "command_bars"
            outText = ToolCommandBars(JsonGetStr(args, "action"), JsonGetStr(args, "bar"), JsonGetStr(args, "control"))
        Case "read_active_dialogs"
            outText = WinOps.ReadActiveDialogsJson(JsonGetStr(args, "cls"), JsonGetBool(args, "include_children", True), JsonGetLong(args, "max"))
        Case "close_active_dialogs"
            outText = WinOps.CloseActiveDialogsJson(JsonGetLong(args, "hwnd"), JsonGetStr(args, "button"), JsonGetBool(args, "dry_run", False), JsonGetBool(args, "allow_host_modal", False))
        Case "send_window_message"
            outText = WinOps.SendWindowMessageJson(JsonGetLong(args, "hwnd"), JsonGetStr(args, "action"), JsonGetStr(args, "message"), JsonGetStr(args, "vk"), JsonGetStr(args, "wparam"), JsonGetStr(args, "lparam"), JsonGetLong(args, "presses"), JsonGetBool(args, "post", False), JsonGetBool(args, "confirm", False))
        Case "quit_ide"
            outText = ToolQuitIde()
        Case "validate_code"
            outText = ToolValidateCode(JsonGetStr(args, "name"), JsonGetStr(args, "code"), JsonGetStr(args, "kind"), JsonGetStr(args, "mode"), JsonGetStr(args, "baseline"))
        Case "check_component"
            outText = ToolCheckComponent(JsonGetStr(args, "name"))
        Case "read_source"
            outText = ToolReadSource(JsonGetStr(args, "path"))
        Case "build"
            outText = ToolBuild(JsonGetStr(args, "project"), JsonGetStr(args, "defines"), JsonGetStr(args, "out_path"), JsonGetLong(args, "timeout_ms"), HasArg(args, "laa"), JsonGetBool(args, "laa", True), False)
        Case "edit_by_anchor"
            outText = ToolEditByAnchor(JsonGetStr(args, "name"), JsonGetStr(args, "anchor"), JsonGetStr(args, "new_text"), JsonGetStr(args, "expected_sha256"), JsonGetLong(args, "occurrence"), JsonGetStr(args, "match"), ScopeObjOf(args, "scope"), JsonGetStr(args, "validate"), SnapshotOptIn(args))
        Case "apply_edits"
            outText = ToolApplyEdits(JsonGetStr(args, "name"), JsonGetAny(args, "edits"), JsonGetStr(args, "expected_sha256"), JsonGetStr(args, "match"), JsonGetStr(args, "validate"), SnapshotOptIn(args))
        Case "preview_edit"
            outText = ToolPreviewEdit(JsonGetStr(args, "name"), JsonGetAny(args, "edits"), JsonGetStr(args, "match"))
        Case "write_source"
            outText = ToolWriteSource(JsonGetStr(args, "path"), JsonGetStr(args, "text"), JsonGetStr(args, "encoding"), JsonGetStr(args, "validate"))
        Case "verify"
            outText = ToolVerify(JsonGetStr(args, "name"), JsonGetStr(args, "disk_path"), JsonGetStr(args, "baseline"))
        Case "inject_trace"
            outText = ToolInjectTrace(JsonGetStr(args, "name"), JsonGetStr(args, "procedure"), JsonGetStr(args, "expressions"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "remove_trace"
            outText = ToolRemoveTrace(JsonGetStr(args, "name"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "fault_inject"
            outText = ToolFaultInject(JsonGetStr(args, "name"), JsonGetStr(args, "procedure"), JsonGetStr(args, "mode"), JsonGetStr(args, "value"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "clear_faults"
            outText = ToolClearFaults(JsonGetStr(args, "name"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "audit_tail"
            outText = ToolAuditTail(JsonGetLong(args, "count"))
        Case "set_breakpoint"
            outText = ToolSetBreakpoint(JsonGetStr(args, "name"), JsonGetStr(args, "anchor"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "clear_breakpoints"
            outText = ToolClearBreakpoints(JsonGetStr(args, "name"), JsonGetStr(args, "id"), JsonGetStr(args, "expected_sha256"))
        Case "debug_commands"
            outText = ToolDebugCommands()
        Case "ide_state"
            outText = ToolIdeState(JsonGetBool(args, "cancel_inflight", False), JsonGetStr(args, "cancel_target"), JsonGetBool(args, "release_inflight", False), JsonGetLong(args, "release_min_held_ms"))
        Case "ide_quiet_mode"
            outText = ToolIdeQuietMode(JsonGetBool(args, "on", False), JsonGetLong(args, "ttl_ms"))
        Case "immediate_eval"
            outText = ToolImmediateEval(JsonGetStr(args, "expression"), JsonGetLong(args, "timeout_ms", 5000))
        Case "immediate_run"
            outText = ToolImmediateRun(JsonGetStr(args, "code"), JsonGetLong(args, "timeout_ms", 5000))
        Case "current_position"
            outText = ToolCurrentPosition()
        Case "list_breakpoints"
            outText = ToolListBreakpoints()
        Case "ui_text"
            outText = WinProbe.UiTextJson(JsonGetLong(args, "hwnd"), JsonGetStr(args, "cls"), JsonGetLong(args, "depth"))
        Case "find_gui"
            outText = ToolFindGui(args)
        Case "execution_state"
            outText = ToolExecutionState(JsonGetLong(args, "pid"))
        Case "vba6_probe_plan"
            outText = Vba6Probe.ProbePlanJson()
        Case "get_callstack"
            outText = Vba6Probe.CallstackJson(JsonGetLong(args, "max_frames", 32))
        Case "vba6_probe"
            If JsonGetBool(args, "hosted_expr_srv", False) Then
                ' NEW DIRECTION: instantiate the engine's hosted expression service
                outText = Vba6Probe.HostedExprSrvJson()
            ElseIf JsonGetBool(args, "wnd_probe", False) Then
                ' RE round-32/33 follow-up: does VBA6.DLL own any window?
                outText = Vba6Probe.WindowProbeJson(JsonGetLong(args, "wnd_max", 200))
            ElseIf JsonGetBool(args, "mem_scan", False) Then
                ' RE round-32: replicate the whole-image scan over LIVE memory
                outText = Vba6Probe.MemScanJson(ParseHexOrLong(JsonGetStr(args, "scan_addr")), JsonGetLong(args, "scan_len"), JsonGetLong(args, "scan_lo"), JsonGetLong(args, "scan_hi"), JsonGetLong(args, "scan_max"))
            ElseIf JsonGetBool(args, "export_rvas", False) Then
                ' RE round-25/26: verify every claimed RVA by name lookup
                outText = Vba6Probe.ExportRvasJson()
            ElseIf JsonGetBool(args, "engine_items", False) Then
                ' RE round-24: the engine's own 72-item capability table
                outText = Vba6Probe.EngineItemsJson()
            ElseIf JsonGetLong(args, "cs_fn", -1) >= 0 Then
                ' RE round-23: the authoritative 5-argument callstack call
                outText = Vba6Probe.CallstackFunctionJson(JsonGetLong(args, "cs_fn", -1))
            ElseIf JsonGetBool(args, "mem_dump", False) Then
                ' RE round-23: "mem_rva" may NAME the image -- "vb6+0x1234" or
                ' "vba6+0x1234". A plain number keeps its historic VBA6.DLL meaning,
                ' so no existing caller changes behaviour.
                Dim r23Spec As String, r23Which As String, r23Rva As Long, r23Mod As Long
                r23Spec = JsonGetStr(args, "mem_rva")
                r23Rva = Vba6Probe.ResolveRvaSpec(r23Spec, r23Which)
                r23Mod = 0
                If r23Which = "vb6" Then r23Mod = Vba6Probe.HostBase()
                outText = Vba6Probe.MemDumpJson(r23Rva, ParseHexOrLong(JsonGetStr(args, "mem_addr")), JsonGetLong(args, "mem_len"), r23Mod)
            ElseIf JsonGetBool(args, "worker_call", False) Then
                outText = Vba6Probe.WorkerCallJson(JsonGetStr(args, "worker_text"), JsonGetBool(args, "worker_bstr", True), JsonGetLong(args, "worker_p1", 1), JsonGetLong(args, "worker_p2", 1), JsonGetLong(args, "worker_p3", 0), JsonGetLong(args, "worker_p4", 0), JsonGetLong(args, "worker_p5", 1), JsonGetLong(args, "worker_p6", 1), JsonGetLong(args, "worker_p7", 0), JsonGetBool(args, "worker_set_bit", False), JsonGetBool(args, "worker_p2_engine", False))
            ElseIf Len(JsonGetStr(args, "rva_call")) > 0 Then
                ' RE round-19: "rva_call" may carry an argument list after ';',
                ' e.g. "0x61D34;0,4,0" -> CallRvaArgsJson (N args, N<=6).
                Dim rcCallSpec As String, rcCallSemi As Long, rcPipe As Long, rcDumpLen As Long
                rcCallSpec = JsonGetStr(args, "rva_call")
                ' RE round-23: a leading "vb6:" targets the HOST image (VB6.EXE),
                ' e.g. "vb6:0x1234;a1,a2". No prefix = VBA6.DLL (unchanged).
                Dim r23Mod2 As Long
                r23Mod2 = 0
                If LCase$(Left$(rcCallSpec, 4)) = "vb6:" Then
                    r23Mod2 = Vba6Probe.HostBase()
                    rcCallSpec = Mid$(rcCallSpec, 5)
                End If
                ' optional "|dump=N" -> atomic create+dump of the returned pointer
                rcDumpLen = 0
                rcPipe = InStr(rcCallSpec, "|dump=")
                If rcPipe > 0 Then
                    rcDumpLen = CLng(Val(Mid$(rcCallSpec, rcPipe + 6)))
                    rcCallSpec = Left$(rcCallSpec, rcPipe - 1)
                End If
                rcCallSemi = InStr(rcCallSpec, ";")
                If rcCallSemi > 0 Then
                    outText = Vba6Probe.CallRvaArgsJson(ParseHexOrLong(Left$(rcCallSpec, rcCallSemi - 1)), Mid$(rcCallSpec, rcCallSemi + 1), rcDumpLen, r23Mod2)
                Else
                    outText = Vba6Probe.CallRvaJson(ParseHexOrLong(rcCallSpec), JsonGetLong(args, "rva_argc"), JsonGetLong(args, "rva_a1"), JsonGetLong(args, "rva_a2"), r23Mod2)
                End If
            ElseIf JsonGetBool(args, "frame_set", False) Then
                outText = Vba6Probe.CallFrameSetterJson(JsonGetLong(args, "frame_set_a1"), JsonGetLong(args, "frame_set_a2", 1))
            ElseIf JsonGetBool(args, "frame_probe", False) Then
                outText = Vba6Probe.FrameStateJson()
            ElseIf JsonGetBool(args, "exec_probe", False) Then
                ' TASK-P1-160 (R17): the exec probe goes through a wrapper that splits
                ' "the statement raised" from "the forward/channel died" and carries the VB
                ' error as structure. The other probe branches are untouched.
                outText = ToolVba6ExecProbe(args)
            Else
                outText = Vba6Probe.ProbeRunJson(JsonGetBool(args, "run_break_probes", False), JsonGetBool(args, "force", False), JsonGetStr(args, "ident"), JsonGetStr(args, "step"), JsonGetLong(args, "frame_index"))
            End If
        Case "project_changes"
            outText = Changes.RevisionJson(JsonGetLong(args, "since", 0), JsonGetBool(args, "clear", False))
        Case "export_public_api"
            outText = ToolExportPublicApi(JsonGetStr(args, "name"))
        Case "check_public_api"
            outText = ToolCheckPublicApi(JsonGetStr(args, "name"), JsonGetStr(args, "baseline"))
        Case "dependency_graph"
            outText = ToolDependencyGraph()
        Case "purity_check"
            outText = ToolPurityCheck(JsonGetStr(args, "name"))
        Case "generate_stub"
            outText = ToolGenerateStub(JsonGetStr(args, "name"))
        Case "record_fixture"
            outText = Fixture.FixRecordJson(ResolvePath(JsonGetStr(args, "path")), JsonGetStr(args, "text"))
        Case "compare_fixture"
            outText = Fixture.FixCompareJson(ResolvePath(JsonGetStr(args, "path")), JsonGetStr(args, "text"))
        Case "list_lessons"
            outText = Lessons.ListJson(JsonGetBool(args, "includeDisabled", False))
        Case "add_lesson"
            outText = Lessons.AddJson(JsonGetStr(args, "id"), JsonGetStr(args, "scenario"), JsonGetStr(args, "trap"), JsonGetStr(args, "do"), JsonGetStr(args, "scope"), JsonGetBool(args, "enabled", True))
        Case "delete_lesson"
            outText = Lessons.DeleteJson(JsonGetStr(args, "id"))
        Case "export_lessons"
            outText = Lessons.ExportMd(JsonGetBool(args, "includeDisabled", False))
        Case "import_lessons_md"
            ' TASK-P1-113 (criterion 3): this used to answer ok:false with ONLY an \`error\` string and
            ' no \`reason\`, so a caller could not branch without parsing prose. The schema now also
            ' DECLARES path as required (registration above), so declaration and enforcement agree.
            ' The verdict vocabulary is the TASK-P1-108 classifier's; \`error\` is kept for callers that
            ' already read it.
            Dim impPath As String
            impPath = JsonGetStr(args, "path")
            If Len(Trim$(impPath)) = 0 Then
                Dim impErr As Object
                Set impErr = CreateObject("Scripting.Dictionary")
                impErr("ok") = False
                impErr("reason") = "required_argument_missing"
                impErr("tool") = "import_lessons_md"
                impErr("argument") = "path"
                impErr("error") = "path is required (the markdown file to import)"
                impErr("message") = "path is required (the markdown file to import)"
                impErr("suggestedAction") = "supply the reported argument (see 'argument' / 'message') and retry"
                impErr("sideEffects") = "none: the call was refused before any file was read or written"
                outText = JsonStringify(impErr)
            Else
                outText = Lessons.ImportMdJson(ResolvePath(impPath), JsonGetBool(args, "dryRun", False))
            End If
        Case "ext_list"
            outText = Extensions.ListJson()
        Case "ext_enable"
            outText = Extensions.EnableJson(JsonGetStr(args, "name"))
        Case "ext_disable"
            outText = Extensions.DisableJson(JsonGetStr(args, "name"))
        Case "ext_reload"
            outText = Extensions.ReloadJson()
        Case "ext_install"
            outText = Extensions.InstallJson(ResolvePath(JsonGetStr(args, "path")))
        Case "ext_uninstall"
            outText = Extensions.UninstallJson(JsonGetStr(args, "name"))
        Case "workspace_status"
            outText = ToolWorkspaceStatus(JsonGetStr(args, "repo"))
        Case "verify_scope"
            outText = ToolVerifyScope(AnyObjOf(args, "components"), JsonGetStr(args, "path"))
        Case "todo_status"
            outText = ToolTodoStatus()
        Case "mcp_info"
            outText = ToolMcpInfo()
        Case "replace_in_files"
            outText = ToolReplaceInFiles(JsonGetStr(args, "path"), JsonGetStr(args, "query"), JsonGetStr(args, "replacement"), _
                JsonGetBool(args, "regex", False), JsonGetBool(args, "dry_run", True), AnyObjOf(args, "expected_shas"), _
                JsonGetStr(args, "validate", "warn"), JsonGetLong(args, "max_files"))
        Case "undo_edit"
            outText = ToolUndoEdit(JsonGetStr(args, "component"), JsonGetStr(args, "action", "restore"), _
                JsonGetStr(args, "ts"), JsonGetBool(args, "to_disk", False))
        Case Else
            ' extension library dispatch (returns empty when the name is not ours)
            outText = Extensions.Dispatch(name, args)
            If Len(outText) = 0 Then Err.Raise vbObjectError + 100, "CodeTools", "Unknown tool: " & name
    End Select
    item("text") = StampMcp(SealedReceipt(name, outText))
    content.Add item
    Set res("content") = content
    res("isError") = False
    Audit.AuditAdd name, "ok", Left$(outText, 200)
    Set CallToolInner = res
    Exit Function
Fail:
    ' never surface an empty message: an empty Err.Description used to come back
    ' as a bare "Error: " that hid the real cause
    Dim failMsg As String
    failMsg = Err.Description
    If Len(failMsg) = 0 Then failMsg = "error " & CStr(Err.Number) & " in " & Err.Source
    ' TASK-P1-039 (contract v1.3.2): capture the error number HERE, once, and clear it.
    ' VB6@s Err object is per-thread and persists; several helpers run between this point
    ' and the JSON builders, so re-reading Err late could attribute a stale or unrelated
    ' code to this receipt. Only this captured value is ever reported.
    Dim failVbErr As Long
    failVbErr = Err.Number
    Err.Clear
    Audit.AuditAdd name, "error", failMsg
    ' TASK-P1-027: a tagged guard is a REFUSAL, not an error - answer it with the same
    ' structured shape as GuardReject (ok:false + reason), not a bare message.
    If Left$(failMsg, 10) = "MCPREASON:" Then
        Dim mcpRes As Object, mcpContent As Object, mcpItem As Object
        Set mcpRes = CreateObject("Scripting.Dictionary")
        Set mcpContent = New Collection
        Set mcpItem = CreateObject("Scripting.Dictionary")
        mcpItem("type") = "text"
        mcpItem("text") = StampMcp(SealedReceipt(name, McpReasonJson(failMsg, name, args)))
        mcpContent.Add mcpItem
        Set mcpRes("content") = mcpContent
        mcpRes("isError") = False
        Set CallToolInner = mcpRes
        Exit Function
    End If
    ' Contract v1.1.0 G-7 (non-schema, TASK-P1-019): while the debugger is stopped the
    ' IDE refuses source writes with a bare localized "permission denied". That is not
    ' a generic failure - the caller must be told WHY and how to recover, so a raw
    ' VB6 permission error in break mode is answered structurally.
    If IsIdeBusyPermissionError(failMsg) Then
        item("text") = StampMcp(SealedReceipt(name, IdeBusyBreakJson(name, failMsg, failVbErr)))
        content.Add item
        Set res("content") = content
        res("isError") = True
        Set CallToolInner = res
        Exit Function
    End If
    ' TASK-P1-108 (D-1107-01): a CLASSIFIED condition (our own raise whose message family
    ' is known, or a measured native VB/COM code) answers with a branchable STRUCTURED
    ' refusal instead of the bare ok:false/reason:"unexpected_error" envelope. This runs
    ' only after the MCPREASON branch and the IDE-busy branch, i.e. only on calls that used
    ' to answer unexpected_error - the success path and every existing structured refusal
    ' never reach here. An empty return keeps the old envelope byte-identical.
    Dim clsRefusalJson As String
    clsRefusalJson = ClassifiedRefusalJson(name, failMsg, failVbErr)
    If Len(clsRefusalJson) > 0 Then
        item("text") = StampMcp(SealedReceipt(name, clsRefusalJson))
        content.Add item
        Set res("content") = content
        res("isError") = False
        Set CallToolInner = res
        Exit Function
    End If
    ' TASK-P1-062: structured envelope for an unexpected VB error (the tagged-guard and
    ' break-mode branches above are unchanged).
    item("text") = StampMcp(SealedReceipt(name, UnexpectedErrorJson(name, failMsg, failVbErr)))
    content.Add item
    Set res("content") = content
    res("isError") = True
    Set CallToolInner = res
End Function

' =====================================================================
' project tools
' =====================================================================

Private Function ToolGetProjectInfo() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' ideState works in every mode (toolbar + caption fallback)
    d("ideState") = GetIdeState()
    ' ---------------------------------------------------------------------------
    ' TASK-P1-090: "no project open" is a LEGAL IDE state, not a fault. It used to fall
    ' out of GetProject() as a raw VB error and came back as ok:false/unexpected_error,
    ' which made the STATE AXIS unusable exactly when a caller needs it (right after
    ' remove_project, with the settings form open, or from automation). The question is
    ' answered instead of failing it: name/fileName are null, the lists are empty and
    ' noProject says so. (The probe P1-079 used read this receipt's MISSING name as
    ' "name: null" - the shape is now explicit.)
    ' ---------------------------------------------------------------------------
    If m_vbe Is Nothing Then
        d("noProject") = True
        d("name") = Null
        d("fileName") = Null
        d("componentCount") = 0
        Set d("components") = New Collection
        Set d("references") = New Collection
        Set d("projects") = New Collection
        d("note") = "the add-in has no live VBIDE connection; nothing could be read"
        ToolGetProjectInfo = JsonStringify(d)
        Exit Function
    End If
    If m_vbe.VBProjects.Count = 0 Then
        d("noProject") = True
        d("name") = Null
        d("fileName") = Null
        d("componentCount") = 0
        Set d("components") = New Collection
        Set d("references") = New Collection
        Set d("projects") = New Collection
        d("note") = "no project is open in the IDE; the project-scoped fields are null/empty"
        ToolGetProjectInfo = JsonStringify(d)
        Exit Function
    End If
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    ' run mode locks most VBProject properties ('Access denied'): read
    ' them all under On Error Resume Next and emit what we can
    On Error Resume Next
    d("name") = p.Name
    d("fileName") = p.FileName
    d("projectType") = ProjectTypeName(p)
    d("description") = p.Description
    d("vbVersion") = m_vbe.Version
    d("vbeName") = m_vbe.Name
    d("vbeFullName") = m_vbe.FullName
    d("templatePath") = m_vbe.TemplatePath
    d("projectSaved") = p.Saved
    d("helpFile") = p.HelpFile
    d("helpContextID") = p.HelpContextID
    d("buildFileName") = p.BuildFileName
    d("compatibleOleServer") = p.CompatibleOleServer
    d("startMode") = StartModeName(p)
    Dim sup As Variant
    sup = p.VBComponents.StartUpObject
    If Not IsNull(sup) And Not IsEmpty(sup) Then d("startupObject") = CStr(sup)
    d("groupFileName") = m_vbe.VBProjects.FileName
    Dim sp As VBIDE.VBProject
    Set sp = m_vbe.VBProjects.StartProject
    If Not sp Is Nothing Then d("startProject") = sp.Name
    d("readOnlyMode") = m_vbe.ReadOnlyMode
    d("displayModel") = DisplayModelName()
    Set d("components") = ComponentListArr()
    ' TASK-P1-090: componentCount exists in BOTH the normal and the no-project receipt, so a
    ' caller does not have to special-case the field; 0 means "no components" in either state.
    d("componentCount") = d("components").Count
    Set d("references") = ReferenceListArr()
    Set d("projects") = ProjectListArr()
    On Error GoTo 0
    ToolGetProjectInfo = JsonStringify(d)
End Function

Private Function ToolSetProjectOptions(ByVal description As String, ByVal helpFile As String, ByVal helpContextId As Long, ByVal startupObject As String, ByVal startProject As String, ByVal buildFileName As String, ByVal activeProject As String, ByVal readOnlyMode As Boolean, ByVal readOnlyModeSet As Boolean) As String
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim applied As Object
    Set applied = New Collection
    On Error GoTo Fail
    If Len(description) > 0 Then
        p.Description = description
        applied.Add "description"
    End If
    If Len(helpFile) > 0 Then
        p.HelpFile = helpFile
        applied.Add "helpFile"
    End If
    If helpContextId <> 0 Then
        p.HelpContextID = helpContextId
        applied.Add "helpContextID"
    End If
    Dim startupObjectError As String
    startupObjectError = ""
    If Len(startupObject) > 0 Then
        Dim vbps0 As VBIDE.VBComponents
        Set vbps0 = p.VBComponents
        ' TASK-P1-023: this build rejects the VBIDE StartUpObject put for EVERY value
        ' (measured: "Sub Main", the module name and the form name all raise
        ' "invalid procedure call or argument"), so report it structurally and let the
        ' other options still apply instead of failing the whole call.
        On Error Resume Next
        Err.Clear
        vbps0.StartUpObject = startupObject
        If Err.Number <> 0 Then
            startupObjectError = Err.Description & " (0x" & Hex$(Err.Number) & ")"
            Err.Clear
        End If
        On Error GoTo Fail
        If Len(startupObjectError) = 0 Then applied.Add "startupObject"
    End If
    If Len(buildFileName) > 0 Then
        p.BuildFileName = buildFileName
        applied.Add "buildFileName"
    End If
    If Len(startProject) > 0 Then
        Dim tgt As VBIDE.VBProject
        Set tgt = FindProject(startProject)
        Dim vbps1 As VBIDE.VBProjects
        Set vbps1 = m_vbe.VBProjects
        vbps1.StartProject = tgt
        applied.Add "startProject"
    End If
    If Len(activeProject) > 0 Then
        Dim ap As VBIDE.VBProject
        Set ap = FindProject(activeProject)
        Set m_vbe.ActiveVBProject = ap
        applied.Add "activeProject"
    End If
    If readOnlyModeSet Then
        m_vbe.ReadOnlyMode = IIf(readOnlyMode, 1, 0)
        applied.Add "readOnlyMode"
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("project") = p.Name
    Set root("applied") = applied
    If Len(startupObjectError) > 0 Then
        root("startupObjectError") = startupObjectError
        root("startupObjectWorkaround") = "the VBIDE StartUpObject put is rejected on this build for every value; the supported route is the project file: set Startup=""<FormName|Sub Main>"" in the .vbp and reopen the project (measured: with Startup=""Sub Main"" a run starts at modMain.Main)"
    End If
    ToolSetProjectOptions = JsonStringify(root)
    Exit Function
Fail:
    Err.Raise vbObjectError + 43, "CodeTools", "set_project_options failed: " & Err.Description
End Function

Private Function ToolOpenProject(ByVal path As String, ByVal exclusive As Boolean) As String
    ' D5 hardening. open_project used to report "type mismatch" (13) and then
    ' leave the MCP endpoint unreachable. Two properties are now measured:
    '   * every step is written and closed on disk BEFORE the risky statement
    '     runs, so a wedged IDE still leaves a log of how far the call got;
    '   * success is decided by the IDE's own project list, never by the call's
    '     return value - the API has been observed to raise AFTER it added.
    ' The 2-argument AddFromFile(PathName, Exclusive) documented in
    ' 内部 VBIDE 对象模型笔记 is tried first, then the 1-argument form.
    Dim steps As Object
    Set steps = New Collection
    OpLog "==== open_project path=" & path & " exclusive=" & CStr(exclusive) & " env(before)=" & OpEnv()
    Dim full As String
    full = ResolvePath(path)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(full) Then Err.Raise vbObjectError + 44, "CodeTools", "Project file not found: " & full
    OpLog "resolved=" & full
    Dim closed As Long
    Dim skippedUnsaved As Long
    Dim apiErr As Long
    Dim apiDesc As String
    closed = 0
    skippedUnsaved = 0
    apiErr = 0
    apiDesc = ""
    Dim alreadyOpen As Boolean
    Dim savedBeforeClose As Boolean
    Dim missingInVbp As String
    ' TASK-P1-165 (R16-1): reloaded/keptInMemory + the disk-vs-memory comparison + the
    ' in-memory project echo. A real reload happens only when nothing unsaved would be
    ' lost; otherwise the request is REFUSED structurally (never silently ignored).
    Dim reloaded As Boolean
    Dim keptInMemory As Boolean
    Dim reloadRefusal As String
    Dim reloadNote As String
    Dim reloadClosed As Boolean
    Dim diskCmp As Object
    Dim diskCmpAfter As Object
    Dim memBefore As Object
    Dim unsavedInfo As Object
    Dim projSaved As Boolean
    Dim savedReadable As Boolean
    Dim mustRefuse As Boolean
    Dim reloadNotAttempted As Boolean
    alreadyOpen = False
    savedBeforeClose = False
    missingInVbp = MissingVbpFiles(full)
    reloaded = False
    keptInMemory = False
    reloadRefusal = ""
    reloadNote = ""
    reloadClosed = False
    reloadNotAttempted = False
    Set diskCmp = Nothing
    Set diskCmpAfter = Nothing
    Set memBefore = Nothing
    Dim p As VBIDE.VBProject
    Set p = FindOpenProject(full)
    If Not p Is Nothing Then
        alreadyOpen = True
        OpLog "already open as " & p.Name
        ' The project on disk is compared with the project the IDE HOLDS. "alreadyOpen"
        ' alone used to be the whole answer, which read as "the file you named was
        ' loaded" - it was not.
        ' Read the dirty flag EXACTLY ONCE and reuse it everywhere (echo + gate).
        projSaved = False
        savedReadable = True
        On Error Resume Next
        Err.Clear
        projSaved = p.Saved
        If Err.Number <> 0 Then
            savedReadable = False
            projSaved = False
            Err.Clear
        End If
        On Error GoTo 0
        OpLog "open_project already-open: VBProject.Saved=" & CStr(projSaved) & " readable=" & CStr(savedReadable)
        Set memBefore = InMemoryProjectProps(p)
        memBefore("saved") = projSaved
        memBefore("savedReadable") = savedReadable
        Set diskCmp = VbpDiskVsMemory(full, p)
        Set unsavedInfo = CreateObject("Scripting.Dictionary")
        keptInMemory = True
        OpLog "diskVsMemory matches=" & CStr(diskCmp("matches"))
        If Not CBool(diskCmp("matches")) Then
            mustRefuse = ProjectHasUnsavedWork(projSaved, savedReadable, p, unsavedInfo)
            OpLog "gate: projSaved=" & CStr(projSaved) & " readable=" & CStr(savedReadable) & " mustRefuse=" & CStr(mustRefuse)
            ' NO RELOAD IS ATTEMPTED. MEASURED 2026-10-13 on this build: the close+re-add
            ' route (VBProjects.Remove, then VBProjects.AddFromFile) DID remove the project and
            ' then raised 59999 "Method '~' of object '~' failed", leaving the workspace with
            ' ZERO projects; the immediately following AddFromFile failed the same way, so the
            ' IDE could not recover inside that call. A tool must not destroy the workspace in
            ' order to refresh it, so a divergence is answered with a STRUCTURED REFUSAL and
            ' nothing is closed. See reloadNotAttemptedNote.
            reloadNotAttempted = True
            If savedReadable = False Then
                reloadRefusal = "project_saved_flag_unreadable"
                reloadNote = "the .vbp on disk does not describe the project the IDE holds, and VBProject.Saved could not be read on this build, so the tool cannot prove that the in-memory copy has nothing unsaved. No reload is attempted (see reloadNotAttemptedNote); the in-memory copy is still the one in use. See unsavedWorkBasis."
            ElseIf mustRefuse = True Then
                reloadRefusal = "unsaved_changes_would_be_lost"
                reloadNote = "the .vbp on disk does not describe the project the IDE holds, and the in-memory project has UNSAVED changes (VBProject.Saved = False, read once and reported in unsavedWorkBasis); no reload is attempted (see reloadNotAttemptedNote). Two ways forward: (a) call save_project and then open_project again - but be aware save_project REWRITES the .vbp, so it will overwrite the disk edits you were trying to pick up; or (b) call remove_project and then open_project to abandon the in-memory copy and load the disk copy."
            Else
                reloadRefusal = "disk_vs_memory_mismatch"
                reloadNote = "the .vbp on disk does not describe the project the IDE holds, and the in-memory copy reports no unsaved changes - but the reload is NOT attempted (see reloadNotAttemptedNote). The in-memory copy is still the one in use and the disk edits were NOT applied. To take the disk copy, call remove_project and then open_project."
            End If
            OpLog "reload refused: reason=" & reloadRefusal
        End If
    Else
        If exclusive Then
            ' VB6 holds one project (or project group) at a time, so an
            ' "exclusive" open has to close the others. A project with unsaved
            ' work cannot be closed without VB6 raising a modal save prompt, and
            ' that prompt blocks the IDE message pump, so flush everything to
            ' disk first (best effort, logged) and only then close.
            Dim i As Long
            Dim pj As VBIDE.VBProject
            Dim needSave As Boolean
            needSave = False
            For i = m_vbe.VBProjects.Count To 1 Step -1
                Set pj = m_vbe.VBProjects.Item(i)
                If Not pj Is Nothing Then
                    If LCase$(pj.FileName) <> LCase$(full) Then
                        If Not pj.Saved Then needSave = True
                    End If
                End If
            Next
            If needSave Then
                OpLog "unsaved project present; flushing before close"
                Dim preSave As String
                preSave = ""
                On Error Resume Next
                preSave = ToolSaveProject("")
                If Err.Number <> 0 Then
                    Err.Clear
                    preSave = "<save raised>"
                End If
                On Error GoTo 0
                OpLog "pre-close save: " & Left$(preSave, 220)
                savedBeforeClose = True
            End If
            For i = m_vbe.VBProjects.Count To 1 Step -1
                Set pj = m_vbe.VBProjects.Item(i)
                If Not pj Is Nothing Then
                    If LCase$(pj.FileName) <> LCase$(full) Then
                        OpLog "closing project " & pj.Name & " file=" & pj.FileName & " saved=" & CStr(pj.Saved)
                        If CloseProject(pj) Then
                            closed = closed + 1
                        Else
                            skippedUnsaved = skippedUnsaved + 1
                            OpLog "close failed (left open)"
                        End If
                    End If
                End If
            Next
        End If
        AuditAdd "open_project", "begin", full
        Dim nBefore As Long
        nBefore = m_vbe.VBProjects.Count
        ' The 1-argument call is tried first: the 2-argument form is the one
        ' documented in 内部 VBIDE 对象模型笔记 but was measured to raise 13
        ' (type mismatch) while still adding the project, so prefer the quieter
        ' call and only fall back to the documented one when it did nothing.
        OpLog "AddFromFile(1-arg) projects=" & CStr(nBefore)
        On Error Resume Next
        Set p = m_vbe.VBProjects.AddFromFile(full)
        apiErr = Err.Number
        apiDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        OpLog "AddFromFile(1-arg) err=" & CStr(apiErr) & " desc=" & apiDesc & " env=" & OpEnv()
        steps.Add "AddFromFile(path): err=" & CStr(apiErr) & " " & apiDesc
        Set p = FindOpenProject(full)
        If p Is Nothing Then
            Dim eB As Long
            Dim dB As String
            OpLog "AddFromFile(2-arg) retry"
            On Error Resume Next
            Set p = m_vbe.VBProjects.AddFromFile(full, exclusive)
            eB = Err.Number
            dB = Err.Description
            Err.Clear
            On Error GoTo 0
            OpLog "AddFromFile(2-arg) err=" & CStr(eB) & " desc=" & dB & " env=" & OpEnv()
            steps.Add "AddFromFile(path,exclusive): err=" & CStr(eB) & " " & dB
            Set p = FindOpenProject(full)
        End If
        Dim nAfter As Long
        nAfter = m_vbe.VBProjects.Count
        OpLog "after projects=" & CStr(nAfter) & " found=" & OpenProjectName(p)
        If p Is Nothing Then
            Dim detail As String
            Dim k As Long
            detail = ""
            For k = 1 To steps.Count
                detail = detail & " | " & CStr(steps(k))
            Next
            If Len(missingInVbp) > 0 Then detail = detail & " | dangling .vbp references: " & missingInVbp
            Err.Raise vbObjectError + 44, "CodeTools", "open_project failed: " & full & " (projects " & CStr(nBefore) & "->" & CStr(nAfter) & ")" & detail & " [log: " & OpLogPath() & "]"
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    If p Is Nothing Then
        root("opened") = "<none>"
        root("fileName") = ""
    Else
        root("opened") = p.Name
        root("fileName") = p.FileName
    End If
    root("exclusive") = exclusive
    root("alreadyOpen") = alreadyOpen
    ' TASK-P1-165 (R16-1): the receipt distinguishes the three states explicitly, echoes
    ' what the IDE holds, and reports whether the .vbp on disk describes it.
    root("reloaded") = reloaded
    root("keptInMemory") = keptInMemory
    If alreadyOpen Then
        If Not memBefore Is Nothing Then Set root("inMemory") = memBefore
        If Not diskCmp Is Nothing Then
            Set root("diskVsMemory") = diskCmp
            If reloaded Then
                If Not diskCmpAfter Is Nothing Then
                    root("diskVsMemoryAfterReloadMatches") = diskCmpAfter("matches")
                    Set root("diskVsMemoryAfterReload") = diskCmpAfter
                End If
            End If
        End If
        root("openProjectSemantics") = "alreadyOpen:true means the IDE already had this .vbp open. reloaded tells you whether it was re-read from disk; keptInMemory tells you whether the in-memory copy is still the one in use. diskVsMemory compares the .vbp on disk with the project the IDE holds."
        If Not unsavedInfo Is Nothing Then Set root("unsavedWorkBasis") = unsavedInfo
        root("reloadNotAttempted") = reloadNotAttempted
        If reloadNotAttempted Then
            root("reloadNotAttemptedNote") = "MEASURED 2026-10-13: closing the already-open project and re-adding it from disk (VBProjects.Remove then AddFromFile) raised 59999 and left the workspace with ZERO projects, so open_project never destroys the open project in order to refresh it. A disk-vs-memory divergence is answered with a structured refusal (ok:false + reason) instead."
        End If
        If Len(reloadRefusal) > 0 Then
            root("ok") = False
            root("reason") = reloadRefusal
            root("reloadRefused") = True
            root("reloadNote") = reloadNote
            root("error") = reloadNote
            root("sideEffects") = "none: nothing was closed, reloaded or written"
            root("suggestedAction") = "save_project (or check the .vbp edits on disk), then call open_project again"
        End If
    End If
    If exclusive Then
        root("closedProjects") = closed
        root("skippedUnsavedProjects") = skippedUnsaved
        root("savedBeforeClose") = savedBeforeClose
    End If
    If Len(missingInVbp) > 0 Then
        root("missingComponentFiles") = missingInVbp
        root("missingNote") = "The .vbp references component files that do not exist on disk; VB6 crashes when such a project is saved, so save_project will refuse until they are restored or the components are removed."
    End If
    If apiErr <> 0 Then
        root("apiError") = "0x" & Hex$(apiErr) & " " & apiDesc
        root("note") = "The IDE project list shows the project open, so the API error above is spurious (the open did succeed). Log: " & OpLogPath()
    End If
    ToolOpenProject = JsonStringify(root)
End Function

' ---- open_project diagnostics -------------------------------------------
' The call here can wedge the IDE's message pump, so the log must be flushed
' to disk BEFORE each risky statement: after a wedge there is no response and
' the log is the only evidence of how far it got.

Private Function OpLogPath() As String
    OpLogPath = Environ$("TEMP") & Chr$(92) & "vb6mcp_openproject.log"
End Function

Private Sub OpLog(ByVal msg As String)
    On Error Resume Next
    Dim fnum As Integer
    fnum = FreeFile
    Open OpLogPath() For Append As #fnum
    Print #fnum, Format$(Now, "hh:nn:ss") & " " & msg
    Close #fnum
    On Error GoTo 0
End Sub

' Top-window / dialog census: tells "the API refused" apart from "the IDE is
' blocked by a modal dialog".
Private Function OpEnv() As String
    Dim pid As Long
    Dim s As String
    s = "win=? dlg=?"
    On Error Resume Next
    pid = WinProbe.CurrentProcessId()
    s = "win=" & CStr(WinProbe.CountTopWindows(pid)) & " dlg=" & CStr(WinProbe.CountDialogs(pid))
    If Err.Number <> 0 Then
        Err.Clear
        s = "win=? dlg=?"
    End If
    On Error GoTo 0
    OpEnv = s
End Function

Private Function FindOpenProject(ByVal full As String) As VBIDE.VBProject
    Dim pj As VBIDE.VBProject
    If Len(full) = 0 Then Exit Function
    On Error Resume Next
    For Each pj In m_vbe.VBProjects
        If Len(pj.FileName) > 0 Then
            If StrComp(pj.FileName, full, vbTextCompare) = 0 Then
                Set FindOpenProject = pj
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

' Text-level companion to MissingComponentFiles: reads the .vbp before it is
' opened and reports component lines whose file is not on disk.
Private Function MissingVbpFiles(ByVal vbpPath As String) As String
    Dim lines() As String
    Dim i As Long
    Dim t As String
    Dim eq As Long
    Dim val As String
    Dim bad As String
    Dim dir As String
    Dim fso As Object
    Dim key As String
    Dim sp As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    bad = ""
    If Not ReadLinesCp(vbpPath, lines) Then
        MissingVbpFiles = ""
        Exit Function
    End If
    dir = fso.GetParentFolderName(vbpPath)
    For i = 0 To UBound(lines)
        t = lines(i)
        eq = InStr(t, "=")
        If eq > 0 Then
            key = LCase$(Trim$(Left$(t, eq - 1)))
            If key = "form" Or key = "module" Or key = "class" Or key = "usercontrol" Or key = "propertypage" Or key = "userdocument" Or key = "relateddoc" Then
                val = Trim$(Mid$(t, eq + 1))
                ' VB6 writes "Module=name; file.bas" / "Class=name; file.cls"
                ' (only Form= carries the bare file name), so the file is what
                ' follows the last semicolon.
                sp = InStrRev(val, ";")
                If sp > 0 Then val = Trim$(Mid$(val, sp + 1))
                If Len(val) > 0 And Len(dir) > 0 Then
                    If Not fso.FileExists(CtlJoinPath(dir, val)) Then bad = bad & val & "; "
                End If
            End If
        End If
    Next
    MissingVbpFiles = bad
End Function

' =============================================================
' TASK-P1-165 (R16): what the IDE HOLDS vs what the .vbp SAYS
' =============================================================

' =============================================================
' TASK-P1-165 (R16): what the IDE HOLDS vs what the .vbp SAYS
' =============================================================

' A .vbp line key that declares a COMPONENT (as opposed to a reference, an object or
' a project option). Same set MissingVbpFiles uses.
Private Function VbpComponentKey(ByVal key As String) As Boolean
    VbpComponentKey = False
    Select Case key
        Case "form", "module", "class", "usercontrol", "propertypage", "userdocument", "designer"
            VbpComponentKey = True
    End Select
End Function

' Both sides of the type comparison are normalised through this table, because the .vbp
' spells the project type as a WORD while the object model hands back a NUMBER.
Private Function NormalizeProjectTypeToken(ByVal s As String) As String
    Dim t As String
    t = LCase$(Trim$(s))
    Select Case t
        Case "0", "exe", "standardexe"
            NormalizeProjectTypeToken = "standardexe"
        Case "1", "oleexe", "activexexe"
            NormalizeProjectTypeToken = "activexexe"
        Case "2", "oledll", "activexdll"
            NormalizeProjectTypeToken = "activexdll"
        Case "3", "olectl", "activexcontrol", "control"
            NormalizeProjectTypeToken = "activexcontrol"
        Case Else
            NormalizeProjectTypeToken = t
    End Select
End Function

' Can closing the in-memory project destroy work?
' The gate uses the VBProject.Saved value that the CALLER read ONCE - never a second read
' inside this helper. MEASURED 2026-10-13: within ONE open_project call the first read of
' VBProject.Saved reported True (it was echoed as inMemory.saved) while the gate path still
' concluded "unsaved", so every divergence refused with "unsaved work" for a project that
' save_project had just written (a misdiagnosis, discipline 39). One read, one value, used
' for BOTH the echo and the decision - no ambiguity left.
' VBComponent.Saved is only REPORTED (on this build it raises, so it is reported as
' unreadable); it is never decisive.
Private Function ProjectHasUnsavedWork(ByVal projSaved As Boolean, ByVal savedReadable As Boolean, _
                                      ByVal p As VBIDE.VBProject, ByRef info As Object) As Boolean
    ProjectHasUnsavedWork = False
    info("projectSaved") = projSaved
    info("projectSavedReadable") = savedReadable
    info("unsavedGateBasis") = "VBProject.Saved, read ONCE by the caller and reused for BOTH the inMemory echo and this gate. VBComponent.Saved is reported when readable and is never decisive."
    Dim comps As Object
    Set comps = New Collection
    Set info("componentSavedFlags") = comps
    Dim compReadable As Boolean
    compReadable = True
    On Error Resume Next
    If p Is Nothing Then
        info("componentSavedReadable") = compReadable
        Exit Function
    End If
    Dim i As Long
    For i = 1 To p.VBComponents.Count
        Dim c As VBIDE.VBComponent
        Set c = Nothing
        Set c = p.VBComponents.Item(i)
        If Not c Is Nothing Then
            Dim e As Object
            Set e = CreateObject("Scripting.Dictionary")
            e("name") = c.Name
            Dim sv As Boolean
            sv = False
            Err.Clear
            sv = c.Saved
            If Err.Number <> 0 Then
                compReadable = False
                Err.Clear
                e("savedReadable") = False
            Else
                e("saved") = sv
            End If
            comps.Add e
        End If
        Err.Clear
    Next
    info("componentSavedReadable") = compReadable
    Err.Clear
    ' EXPLICIT comparisons, never 'Not <boolean>': a bitwise Not on a non-Boolean variant
    ' (VB's Not is bitwise when the operand is numeric) would turn 1 into -2, i.e. TRUE,
    ' which is exactly how this gate refused a project that VBProject.Saved said was clean.
    If savedReadable = False Then
        info("unsavedRefusalCause") = "project_saved_flag_unreadable"
        ProjectHasUnsavedWork = True
        info("gateValue") = True
        Exit Function
    End If
    If projSaved = False Then
        info("unsavedRefusalCause") = "project_reports_unsaved"
        ProjectHasUnsavedWork = True
        info("gateValue") = True
        Exit Function
    End If
    ProjectHasUnsavedWork = False
    info("gateValue") = False
    info("gateBasis") = "projSaved=" & CStr(projSaved) & " savedReadable=" & CStr(savedReadable) & " => reload is safe"
End Function

' The properties the VBIDE object model exposes for the OPEN project.
Private Function InMemoryProjectProps(ByVal p As VBIDE.VBProject) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set InMemoryProjectProps = d
    On Error Resume Next
    If p Is Nothing Then Exit Function
    d("name") = p.Name
    d("fileName") = p.FileName
    d("saved") = p.Saved
    d("description") = p.Description
    d("helpFile") = p.HelpFile
    d("helpContextID") = p.HelpContextID
    d("buildFileName") = p.BuildFileName
    d("readOnly") = p.ReadOnly
    d("componentCount") = p.VBComponents.Count
    Dim names As Object
    Set names = New Collection
    Dim i As Long
    For i = 1 To p.VBComponents.Count
        Dim c As VBIDE.VBComponent
        Set c = p.VBComponents.Item(i)
        If Not c Is Nothing Then names.Add c.Name
        Err.Clear
    Next
    Set d("componentNames") = names
    Dim so As String
    so = ""
    so = p.StartUpObject
    If Err.Number <> 0 Then
        Err.Clear
        d("startupObjectUnreadable") = True
    Else
        d("startupObject") = so
    End If
    Err.Clear
End Function

' Compare the .vbp on disk with the project the IDE holds. matches=False is what turns
' "alreadyOpen" into a real reload or a structured refusal.
Private Function VbpDiskVsMemory(ByVal path As String, ByVal p As VBIDE.VBProject) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("matches") = False
    d("diskVbpPath") = path
    d("diskVbpSha256") = ""
    d("diskProjectType") = ""
    d("memoryProjectType") = ""
    d("diskComponentCount") = 0
    d("memoryComponentCount") = 0
    d("decisiveScalarDiffs") = 0
    d("matchesBasis") = "the .vbp's declared COMPONENT FILES and Type= against the components and type the IDE holds; the ExeName32/BuildFileName and Startup/StartUpObject pairs count only when BOTH sides are non-empty (an unset property on one side means 'not reported here', not 'different')"
    Set d("componentsOnlyOnDisk") = New Collection
    Set d("componentsOnlyInMemory") = New Collection
    Set d("scalarDiffs") = New Collection
    Set VbpDiskVsMemory = d
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    If Dir$(path) = "" Then Exit Function
    Dim txt As String
    txt = Build.ReadLogAnsi(path)
    d("diskVbpSha256") = Sha256.Sha256HexTextUtf8(txt)
    Dim lines() As String, i As Long, ln As String, key As String, val As String, eq As Long, sp As Long
    Dim diskSet As Object, memSet As Object
    Set diskSet = CreateObject("Scripting.Dictionary")
    Set memSet = CreateObject("Scripting.Dictionary")
    Dim diskType As String, diskExe As String, diskStartup As String
    diskType = ""
    diskExe = ""
    diskStartup = ""
    lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        eq = InStr(ln, "=")
        If eq > 1 Then
            key = LCase$(Trim$(Left$(ln, eq - 1)))
            val = Trim$(Mid$(ln, eq + 1))
            If key = "type" Then diskType = val
            If key = "exename32" Then diskExe = val
            If key = "startup" Then diskStartup = val
            If VbpComponentKey(key) Then
                sp = InStrRev(val, ";")
                If sp > 0 Then val = Trim$(Mid$(val, sp + 1))
                If Len(val) > 0 Then diskSet(LCase$(BaseNameOf(val))) = True
            End If
        End If
    Next
    d("diskProjectType") = diskType
    d("diskComponentCount") = diskSet.Count
    Dim memType As String, memExe As String, memStartup As String
    memType = ""
    memExe = ""
    memStartup = ""
    If Not p Is Nothing Then
        memType = CStr(p.Type)
        memExe = p.BuildFileName
        memStartup = ""
        memStartup = p.StartUpObject
        If Err.Number <> 0 Then Err.Clear
        For i = 1 To p.VBComponents.Count
            Dim c As VBIDE.VBComponent
            Set c = p.VBComponents.Item(i)
            If Not c Is Nothing Then
                Dim fn As String
                fn = ""
                fn = c.FileNames(1)
                If Len(fn) > 0 Then memSet(LCase$(BaseNameOf(fn))) = True
            End If
            Err.Clear
        Next
    End If
    d("memoryProjectType") = memType
    d("memoryComponentCount") = memSet.Count
    d("typeMatches") = (NormalizeProjectTypeToken(diskType) = NormalizeProjectTypeToken(memType))
    Dim k As Variant
    For Each k In diskSet.Keys
        If Not memSet.Exists(CStr(k)) Then d("componentsOnlyOnDisk").Add CStr(k)
    Next
    For Each k In memSet.Keys
        If Not diskSet.Exists(CStr(k)) Then d("componentsOnlyInMemory").Add CStr(k)
    Next
    AddScalarDiff d, "ExeName32 vs BuildFileName", diskExe, memExe
    AddScalarDiff d, "Startup vs StartUpObject", diskStartup, memStartup
    d("matches") = (d("componentsOnlyOnDisk").Count = 0 And d("componentsOnlyInMemory").Count = 0 And CBool(d("typeMatches")) And d("decisiveScalarDiffs") = 0)
    If Err.Number <> 0 Then Err.Clear
End Function

' One scalar pair. An EMPTY side means "not reported here" and is skipped, so an unset
' BuildFileName can never by itself trigger a reload.
Private Sub AddScalarDiff(ByVal d As Object, ByVal pair As String, ByVal diskV As String, ByVal memV As String)
    Dim dv As String, mv As String
    dv = Trim$(diskV)
    mv = Trim$(memV)
    If Len(dv) = 0 Or Len(mv) = 0 Then Exit Sub
    If StrComp(BaseNameOf(dv), BaseNameOf(mv), vbTextCompare) = 0 Then Exit Sub
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    e("pair") = pair
    e("disk") = dv
    e("memory") = mv
    e("decisive") = True
    d("scalarDiffs").Add e
    d("decisiveScalarDiffs") = d("decisiveScalarDiffs") + 1
End Sub

' TASK-P1-162 (R7-4): the /d-related receipt block, used on BOTH the cache-hit path and
' the real build path so a caller can always see exactly what was passed.
Private Sub CcArgsInto(ByVal root As Object, ByVal cmdArgs As String, ByVal ccExtra As String, _
                       ByVal ccArgs As Object, ByVal defines As String, ByVal useIde As Boolean)
    root("commandLine") = cmdArgs
    root("commandLineApplies") = (Not useIde)
    If useIde Then
        root("commandLineNote") = "this receipt came from the IN-IDE Make route, which starts NO compiler subprocess; the command line shown is the one the isolated /make route would have used (and did NOT run here)."
    Else
        root("commandLineNote") = "the exact command line handed to VB6.EXE (/make ...), including the /d arguments exactly as they were written."
    End If
    root("conditionalCompileArgsRaw") = defines
    root("conditionalCompileArgsExtra") = ccExtra
    root("conditionalCompileArgsCount") = ccArgs.Count
    Set root("conditionalCompileArgs") = ccArgs
    root("conditionalCompileQuoteRule") = "a /d value is quoted ONLY when it contains a space; an unconditional quote folds the quoted text into the macro NAME on this compiler, so the macro silently never applies while the build still reports success."
End Sub


' =====================================================================
' TASK-P1-160 (R17, narrowed): split "the statement raised" from "the forward/channel
' died", and carry the VB error as a STRUCTURED field instead of prose.
'   * exec_error - the add-in DID answer: either the call into the probe raised, or
'                  EbExecuteLine reported a VB error. The number was ALREADY captured
'                  (Vba6Probe.bas vbaError/vbaErrorDesc); it is re-labelled, not re-derived.
'   * forward_timeout / forward_failed - produced by the HUB (bridge/vb6mcp_hub.mjs), not
'                  by the add-in: they mean the add-in never answered at all, so no
'                  add-in receipt - and therefore no exec_error - can exist for them.
' =====================================================================
Private Function ToolVba6ExecProbe(ByVal args As Object) As String
    Dim raw As String
    Dim raised As Boolean
    Dim rNum As Long
    Dim rDesc As String
    raw = ""
    raised = False
    rNum = 0
    rDesc = ""
    On Error Resume Next
    Err.Clear
    raw = Vba6Probe.ExecLineJson(JsonGetStr(args, "exec_line"), JsonGetLong(args, "exec_a2"), JsonGetLong(args, "exec_a3"), JsonGetLong(args, "exec_a4"), JsonGetBool(args, "exec_scope", False), JsonGetBool(args, "exec_dump_ctx", False), JsonGetBool(args, "exec_set_frame", False))
    If Err.Number <> 0 Then
        raised = True
        rNum = Err.Number
        rDesc = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    Dim transportNote As String
    transportNote = "exec_error means the add-in DID answer: either the call into the probe raised, or EbExecuteLine reported a VB error. It is a DIFFERENT failure from forward_timeout / forward_failed, which the HUB reports when the add-in cannot answer at all (the channel is hung, or the call outlived the forward budget) - for those there is no add-in receipt, and therefore no exec_error."
    If raised Or Len(raw) = 0 Then
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("ok") = False
        d("reason") = "exec_error"
        d("tool") = "vba6_probe"
        d("probe") = "exec_probe"
        Dim e As Object
        Set e = CreateObject("Scripting.Dictionary")
        e("vbaErrNumber") = rNum
        e("description") = rDesc
        e("raisedIn") = "the add-in, while calling Vba6Probe.ExecLineJson"
        If Len(raw) = 0 Then e("emptyReceipt") = True
        Set d("exec_error") = e
        d("raw") = Left$(raw, 400)
        d("transportReasonSpace") = transportNote
        ToolVba6ExecProbe = JsonStringify(d)
        Exit Function
    End If
    ' The probe answered. When it captured a VB error while executing the line, re-label the
    ' SAME numbers as a structured exec_error. A clean probe is returned BYTE-IDENTICAL to
    ' before, so no existing successful caller can change behaviour.
    Dim n As Long
    Dim dr As String
    n = 0
    dr = ""
    Dim f1 As String
    f1 = JsonFieldRaw(raw, "vbaError")
    If Len(f1) > 0 Then n = CLng(Val(f1))
    If n = 0 Then
        ToolVba6ExecProbe = raw
        Exit Function
    End If
    dr = JsonFieldRaw(raw, "vbaErrorDesc")
    Dim addd As Object
    Set addd = CreateObject("Scripting.Dictionary")
    addd("ok") = False
    addd("reason") = "exec_error"
    addd("transportReasonSpace") = transportNote
    Dim e2 As Object
    Set e2 = CreateObject("Scripting.Dictionary")
    e2("vbaErrNumber") = n
    e2("description") = dr
    e2("raisedIn") = "Vba6Probe.ExecLineJson (the probe answered with vbaError/vbaErrorDesc)"
    Set addd("exec_error") = e2
    Dim whole As String
    whole = JsonStringify(addd)
    Dim frag As String
    frag = Mid$(whole, 2, Len(whole) - 2)
    Dim inj As String
    inj = JsonPrependMembers(raw, frag)
    If inj = raw Then
        ' The shape was not a plain non-empty JSON object. Say so instead of pretending.
        Dim d3 As Object
        Set d3 = CreateObject("Scripting.Dictionary")
        d3("ok") = False
        d3("reason") = "exec_error"
        d3("tool") = "vba6_probe"
        d3("probe") = "exec_probe"
        Dim e3 As Object
        Set e3 = CreateObject("Scripting.Dictionary")
        e3("vbaErrNumber") = n
        e3("description") = dr
        e3("raisedIn") = "Vba6Probe.ExecLineJson (the probe answered with vbaError)"
        Set d3("exec_error") = e3
        d3("raw") = raw
        d3("relabelFailed") = "the probe's answer was not a plain non-empty JSON object, so it could not be re-labelled in place"
        d3("transportReasonSpace") = transportNote
        ToolVba6ExecProbe = JsonStringify(d3)
        Exit Function
    End If
    ToolVba6ExecProbe = inj
End Function

' =====================================================================
' TASK-P1-168 (B): immediate_eval / immediate_run - run a line in the IDE's own
' immediate context and CARRY THE RESULT BACK MACHINE-READABLY.
'   The caller must NOT have to write Open/Print/Close itself, and the VBA6 metadata has
'   NO channel that returns EbExecuteLine's value. The sink below was measured on a LIVE
'   instance BEFORE this code existed (raw receipts: docs/evidence/TASK-P1-168/):
'     * Open "<p>" For Output As #511: Print #511, 1+1: Close #511   -> ret=0, sink " 2 "
'     * v511 = 3/4: Print #511, TypeName(v511) & "|" & v511: Close #511 -> "Double|0.75"
'     * a leading Close #511 is harmless and makes the handle idempotent
'     * an embedded CRLF makes EbExecuteLine REFUSE the line with ret=40036
'     * On Error Resume Next (ALONE or colon-joined) is REFUSED with ret=40036, so a line
'       that RAISES cannot be made error-tolerant from here. That is the exact mechanism
'       behind defect #10: the IDE's own modal run-time-error box keeps the call alive.
'   Stated instead of hidden:
'     * a raising expression still blocks the IDE (same family as #10). The armed release
'       and ide_state{release_inflight:true} are the RECOVERY, not a hard timeout;
'     * the add-in cannot preempt a statement, so timeout_ms is a GATE-AGE threshold
'       (timeoutMsEnforcedBy says exactly that, hardTimeoutEnforced is false);
'     * the sink is a FILE, so the value is captured only when the statement completes.
'   NOTE: IM_SINK_MAX_BYTES / IM_FILE_NUM are declared in the DECLARATIONS SECTION at the
'   top of this module - VB6 rejects a module-level Const that appears after the first
'   procedure (measured this round: the compiler then reported the unrelated call site of
'   ToolFaultInject as "Sub or Function not defined").
' =====================================================================
Private Function ImSinkDir() As String
    Dim p As String
    p = Util.TempDir() & "\vb6mcp_im"
    On Error Resume Next
    If Len(Dir$(p, vbDirectory)) = 0 Then MkDir p
    On Error GoTo 0
    ImSinkDir = p
End Function

Private Function ImSinkPath(ByVal tag As String) As String
    ImSinkPath = ImSinkDir() & "\" & tag & "_" & CStr(m_callSeq) & "_" & CStr(GetTickCount()) & ".tmp"
End Function

' The immediate context accepts ONE physical line: a line carrying CR or LF is refused
' with ret=40036 (measured). Multi-statement input is therefore joined with ": " - the
' same thing a human would type.
Private Function ImFlatLine(ByVal code As String) As String
    Dim s As String
    s = Replace(code, vbCrLf, ": ")
    s = Replace(s, vbLf, ": ")
    s = Replace(s, vbCr, ": ")
    ImFlatLine = Trim$(s)
End Function

' Redirect the caller's own Debug.Print into the sink. Only a STATEMENT-BOUNDARY
' Debug.Print is rewritten (start of the line or right after a colon) so ordinary text is
' left alone; the number of rewrites is reported rather than assumed.
Private Function ImRewriteDebugPrint(ByVal code As String, ByRef rewrites As Long) As String
    rewrites = 0
    ImRewriteDebugPrint = code
    If Len(code) = 0 Then Exit Function
    Dim rx As Object
    Dim col As Object
    On Error Resume Next
    Set rx = CreateObject("VBScript.RegExp")
    If Err.Number <> 0 Or rx Is Nothing Then
        Err.Clear
        Exit Function
    End If
    rx.Global = True
    rx.IgnoreCase = True
    rx.Pattern = "(^|:)\s*Debug\s*\.\s*Print\b"
    Set col = rx.Execute(code)
    If Err.Number = 0 Then
        If Not col Is Nothing Then rewrites = col.Count
    End If
    Err.Clear
    ImRewriteDebugPrint = rx.Replace(code, "$1 Print #" & CStr(IM_FILE_NUM) & ",")
    If Err.Number <> 0 Then
        ImRewriteDebugPrint = code
        rewrites = 0
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function ImReadSink(ByVal p As String, ByRef wasTruncated As Boolean) As String
    Dim h As Integer
    Dim n As Long
    Dim s As String
    ImReadSink = ""
    wasTruncated = False
    On Error Resume Next
    If Len(p) = 0 Then Exit Function
    If Len(Dir$(p)) = 0 Then Exit Function
    n = FileLen(p)
    If n <= 0 Then Exit Function
    If n > IM_SINK_MAX_BYTES Then
        wasTruncated = True
        n = IM_SINK_MAX_BYTES
    End If
    h = FreeFile
    Open p For Binary As #h
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    s = Space$(n)
    Get #h, , s
    Close #h
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    ImReadSink = s
    Err.Clear
End Function

Private Function ImDeleteSink(ByVal p As String) As Boolean
    ImDeleteSink = False
    On Error Resume Next
    If Len(Dir$(p)) > 0 Then Kill p
    ImDeleteSink = (Len(Dir$(p)) = 0)
    Err.Clear
End Function

Private Function ImJoinLines(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    ImJoinLines = t
End Function

Private Function ImTrimLines(ByVal s As String) As String
    Dim parts() As String
    Dim i As Long
    Dim out As String
    out = ""
    parts = Split(ImJoinLines(s), vbLf)
    For i = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(i))) > 0 Then
            If Len(out) > 0 Then out = out & vbLf
            out = out & Trim$(parts(i))
        End If
    Next i
    ImTrimLines = out
End Function

Private Function ImCountLines(ByVal s As String) As Long
    Dim t As String
    If Len(s) = 0 Then Exit Function
    t = ImJoinLines(s)
    If Right$(t, 1) = vbLf Then t = Left$(t, Len(t) - 1)
    If Len(t) = 0 Then Exit Function
    ImCountLines = UBound(Split(t, vbLf)) + 1
End Function

Private Function ToolImmediateEval(ByVal expression As String, ByVal timeoutMs As Long) As String
    Dim o As Object
    Set o = ImExecuteImmediate("eval", expression, timeoutMs)
    o("expression") = expression
    ToolImmediateEval = JsonStringify(o)
End Function

Private Function ToolImmediateRun(ByVal code As String, ByVal timeoutMs As Long) As String
    Dim o As Object
    Set o = ImExecuteImmediate("run", code, timeoutMs)
    o("code") = code
    ToolImmediateRun = JsonStringify(o)
End Function

Private Function ImExecuteImmediate(ByVal mode As String, ByVal payload As String, ByVal timeoutMs As Long) As Object
    Dim o As Object
    Set o = CreateObject("Scripting.Dictionary")
    o("tool") = IIf(mode = "eval", "immediate_eval", "immediate_run")
    o("mode") = mode
    o("ok") = False
    o("reason") = ""
    o("valueParsed") = False
    o("value") = ""
    o("typeName") = ""
    o("output") = ""
    o("outputTrimmed") = ""
    o("outputBytes") = 0
    o("outputTruncated") = False
    o("outputLines") = 0
    o("debugPrintRewrites") = 0
    o("sinkFile") = ""
    o("sinkDeleted") = False
    o("line") = ""
    o("elapsedMs") = 0
    o("execAnswered") = False
    o("ret") = 0
    o("retHex") = ""
    o("timeoutMs") = timeoutMs
    o("timeoutMsEnforcedBy") = "gate_age_threshold"
    o("hardTimeoutEnforced") = False
    o("statementInterrupted") = False
    o("armedReleaseMs") = timeoutMs
    o("recoveryAction") = "ide_state{release_inflight:true, release_min_held_ms:<timeout_ms>}"
    o("recoveryHint") = "the add-in cannot preempt a statement, so there is NO hard timeout here. If this call never returns, the ARMED release frees the single-flight gate on the next tool call; if the IDE is waiting on its own modal run-time-error box, read_active_dialogs shows it and close_active_dialogs dismisses it."
    o("note") = ""
    If timeoutMs <= 0 Then timeoutMs = 5000
    o("timeoutMs") = timeoutMs
    o("armedReleaseMs") = timeoutMs
    If Len(payload) = 0 Then
        o("reason") = "empty_input"
        o("note") = "pass a non-empty " & IIf(mode = "eval", "expression", "code") & "; nothing was executed"
        Set ImExecuteImmediate = o
        Exit Function
    End If
    Dim sink As String
    sink = ImSinkPath(mode)
    o("sinkFile") = sink
    On Error Resume Next
    If Len(Dir$(sink)) > 0 Then Kill sink
    On Error GoTo 0
    Dim flat As String
    Dim rewrites As Long
    Dim body As String
    flat = ImFlatLine(payload)
    rewrites = 0
    If mode = "eval" Then
        ' Evaluate ONCE into a Variant, then print "TypeName|value". The leading Close #511
        ' makes the handle idempotent even if an earlier line died holding it.
        body = "v511 = " & flat & ": Print #" & CStr(IM_FILE_NUM) & ", TypeName(v511) & ""|"" & v511"
    Else
        body = ImRewriteDebugPrint(flat, rewrites)
        o("debugPrintRewrites") = rewrites
    End If
    Dim execLine As String
    execLine = "Close #" & CStr(IM_FILE_NUM) & ": Open """ & sink & """ For Output As #" & CStr(IM_FILE_NUM) & ": " & body & ": Close #" & CStr(IM_FILE_NUM)
    o("line") = execLine
    ' ---- arm the self-healing release, THEN run
    m_armTool = o("tool")
    m_armMs = timeoutMs
    m_armAt = GetTickCount()
    Dim t0 As Long
    t0 = GetTickCount()
    Dim raw As String
    Dim raised As Boolean
    Dim rNum As Long
    Dim rDesc As String
    raw = ""
    raised = False
    rNum = 0
    rDesc = ""
    On Error Resume Next
    Err.Clear
    raw = Vba6Probe.ExecLineJson(execLine, 0, 0, 0, False, False, False)
    If Err.Number <> 0 Then
        raised = True
        rNum = Err.Number
        rDesc = Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    o("elapsedMs") = GetTickCount() - t0
    ' we came back, so the arm is moot
    m_armMs = 0
    m_armTool = ""
    m_armAt = 0
    If raised Or Len(raw) = 0 Then
        o("reason") = "exec_error"
        o("execError") = "vbErrNumber=" & CStr(rNum) & "; description=" & rDesc
        o("note") = "the add-in DID answer, but the call into the exec probe raised: " & rDesc
        o("raw") = Left$(raw, 400)
    Else
        o("execAnswered") = True
        o("ret") = CLng(Val(JsonFieldRaw(raw, "ret")))
        o("retHex") = JsonFieldRaw(raw, "retHex")
        If o("ret") <> 0 Then
            o("reason") = "engine_refused_line"
            o("note") = "EbExecuteLine returned " & CStr(o("retHex")) & " (ret=" & CStr(o("ret")) & "), so the line was NOT executed. On this channel ret=40036 (0x9C64) means the line was REFUSED - measured for an embedded CRLF and for any line carrying On Error Resume Next."
        End If
    End If
    ' ---- read the sink back (the ONLY result channel)
    Dim txt As String
    Dim trunc As Boolean
    trunc = False
    txt = ImReadSink(sink, trunc)
    o("outputBytes") = Len(txt)
    o("outputTruncated") = trunc
    o("output") = ImJoinLines(txt)
    o("outputTrimmed") = ImTrimLines(txt)
    o("outputLines") = ImCountLines(txt)
    o("sinkDeleted") = ImDeleteSink(sink)
    If mode = "eval" Then
        Dim joined As String
        Dim firstLine As String
        Dim nl As Long
        Dim bar As Long
        joined = ImJoinLines(txt)
        firstLine = joined
        nl = InStr(joined, vbLf)
        If nl > 0 Then firstLine = Left$(joined, nl - 1)
        bar = InStr(firstLine, "|")
        If bar > 0 Then
            o("typeName") = Left$(firstLine, bar - 1)
            o("value") = Mid$(firstLine, bar + 1)
            o("valueParsed") = True
        End If
    End If
    If o("reason") = "" Then
        If mode = "eval" Then
            If o("valueParsed") Then
                o("ok") = True
            Else
                o("reason") = "value_not_parsed"
                o("note") = "the statement completed (ret=0) but the sink carried no parseable TypeName|value line. output/outputTrimmed hold the RAW sink text - nothing is guessed."
            End If
        Else
            o("ok") = True
        End If
    End If
    If o("ok") Then
        If mode = "eval" Then
            o("note") = "the expression was evaluated ONCE in the ACTIVE project's immediate context; the value was carried out through a temp file, so the caller never had to write Open/Print/Close. VB's Print formatting applies (a positive number is written with a leading space)."
        Else
            o("note") = "the code ran in the ACTIVE project's immediate context with each statement-boundary Debug.Print redirected into a temp file the tool reads back; only what THIS call printed is returned. The ? shorthand and _ line continuations are NOT supported; CR/LF in the input is joined with a colon and a space."
        End If
    End If
    Set ImExecuteImmediate = o
End Function

' =====================================================================
' TASK-P1-164 (A): ide_quiet_mode - a session-level switch that stops the add-in from
' STEALING the user's foreground while an agent drives the IDE.
' =====================================================================
Private Function ToolIdeQuietMode(ByVal enable As Boolean, ByVal ttlMs As Long) As String
    Dim fgB As Long
    Dim fgA As Long
    fgB = WoFrontHwnd()
    Dim wasOn As Boolean
    wasOn = QuietActive()
    If enable Then
        If ttlMs <= 0 Then ttlMs = QUIET_TTL_DEFAULT_MS
        m_quietOn = True
        m_quietSince = GetTickCount()
        m_quietTtlMs = ttlMs
    Else
        m_quietOn = False
        m_quietSince = 0
        m_quietTtlMs = 0
    End If
    Dim nowOn As Boolean
    nowOn = QuietActive()
    fgA = WoFrontHwnd()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("tool") = "ide_quiet_mode"
    d("quietMode") = nowOn
    d("previousQuietMode") = wasOn
    d("changed") = (nowOn <> wasOn)
    d("since") = m_quietSince
    d("ttlMs") = m_quietTtlMs
    d("quietExpiresAt") = IIf(nowOn, m_quietSince + m_quietTtlMs, 0)
    d("expiresInMs") = IIf(nowOn, m_quietTtlMs - (GetTickCount() - m_quietSince), 0)
    d("suppressedFocusCount") = m_quietSuppressed
    d("foregroundBefore") = fgB
    d("foregroundAfter") = fgA
    d("foregroundTookFocus") = (fgA <> fgB)
    d("foregroundReadable") = (fgB <> 0 Or fgA <> 0)
    d("advisory") = IIf(nowOn, QUIET_ADVISORY, "")
    d("suppressedPoints") = "control_window{action:""activate""} skips Window.SetFocus while quiet mode is ON (that receipt carries quietModeSuppressedFocus:true). native_breakpoint CANNOT be issued without activating the IDE - it is a Debug-menu command - so quiet mode can only give the foreground BACK there (quietModeForegroundRestored)."
    d("note") = "quiet OFF restores the pre-v1.9.0 behaviour byte for byte: no SetFocus is suppressed and no quiet{} block is added to any receipt. foregroundReadable=false means GetForegroundWindow() returned 0 in this window station (it does here) - that must NOT be read as the focus was left alone."
    ToolIdeQuietMode = JsonStringify(d)
End Function

' TASK-P1-160 (R17): insert extra top-level members into a JSON OBJECT text without
' re-parsing it. Returns jsonText unchanged when the shape is not a plain non-empty
' object, so the caller can detect the refusal with a plain equality test.
Private Function JsonPrependMembers(ByVal jsonText As String, ByVal frag As String) As String
    JsonPrependMembers = jsonText
    If Len(jsonText) = 0 Then Exit Function
    If Len(frag) = 0 Then Exit Function
    Dim s As String
    Dim i As Long
    s = Trim$(jsonText)
    If Len(s) < 2 Then Exit Function
    If Left$(s, 1) <> "{" Then Exit Function
    If Mid$(s, 2, 1) = "}" Then Exit Function
    i = InStr(jsonText, "{")
    If i = 0 Then Exit Function
    JsonPrependMembers = Left$(jsonText, i) & frag & "," & Mid$(jsonText, i + 1)
End Function

Private Function OpenProjectName(ByVal p As VBIDE.VBProject) As String
    OpenProjectName = "<none>"
    On Error Resume Next
    If Not p Is Nothing Then OpenProjectName = p.Name
    If Err.Number <> 0 Then
        Err.Clear
        OpenProjectName = "<none>"
    End If
    On Error GoTo 0
End Function

' VBProjects.Remove is the documented member; VBProject.Close is the fallback.
Private Function CloseProject(ByVal pj As VBIDE.VBProject) As Boolean
    Dim ok As Boolean
    ok = False
    On Error Resume Next
    m_vbe.VBProjects.Remove pj
    If Err.Number = 0 Then ok = True Else Err.Clear
    On Error GoTo 0
    If Not ok Then
        On Error Resume Next
        pj.Close
        If Err.Number = 0 Then ok = True Else Err.Clear
        On Error GoTo 0
    End If
    CloseProject = ok
End Function

' Lists components whose recorded file name does not exist on disk. VB6
' crashes (AV in USER32) when such a project is saved, so the save path refuses
' to run while this list is non-empty.
Private Function MissingComponentFiles(ByVal p As VBIDE.VBProject) As String
    Dim comp As VBIDE.VBComponent
    Dim f As String
    Dim bad As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    bad = ""
    On Error Resume Next
    For Each comp In p.VBComponents
        f = ""
        Err.Clear
        f = comp.FileNames(1)
        If Err.Number <> 0 Then
            Err.Clear
            f = ""
        End If
        If Len(f) > 0 Then
            If Not fso.FileExists(f) Then bad = bad & comp.Name & " -> " & f & "; "
        End If
    Next
    On Error GoTo 0
    MissingComponentFiles = bad
End Function

Private Function ToolAddProject(ByVal projType As String, ByVal templatePath As String) As String
    Dim p As VBIDE.VBProject
    If Len(templatePath) > 0 Then
        Set p = m_vbe.VBProjects.AddFromTemplate(ResolvePath(templatePath), True)
        CloseStartupDialog
    Else
        Dim pt As vbext_ProjectType
        Select Case LCase$(projType)
            Case "standardexe", ""
                pt = vbext_pt_StandardExe
            Case "activexexe"
                pt = vbext_pt_ActiveXExe
            Case "activexdll"
                pt = vbext_pt_ActiveXDll
            Case "activexcontrol"
                pt = vbext_pt_ActiveXControl
            Case Else
                Err.Raise vbObjectError + 45, "CodeTools", "Unknown project type: " & projType
        End Select
        ' Exclusive=True: the new project becomes THE project and the
        ' IDE leaves its startup group mode (running a group prompts
        ' 'Save Project Group As' before every run)
        Set p = m_vbe.VBProjects.Add(pt, True)
        CloseStartupDialog
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("created") = p.Name
    root("type") = ProjectTypeName(p)
    ToolAddProject = JsonStringify(root)
End Function

Private Sub CloseStartupDialog()
    ' dismiss the IDE's startup New-Project dialog (新建工程); this
    ' build keeps it open as a modal window that blocks run/make even
    ' after VBProjects.Add
    ' Card TASK-P1-007 (i18n): this dialog used to be found by its LOCALIZED
    ' caption, so on a VB6 in any language other than Chinese it was never
    ' dismissed and run/make stayed blocked behind a modal window. Match by window
    ' CLASS instead: it is always #32770 and always owned by THIS process. To avoid
    ' touching unrelated dialogs we only act on the one that is currently the
    ' foreground window (a modal dialog is the active window). The historical
    ' caption lookup is kept ONLY as a last resort - nothing depends on it.
    Const WM_CLOSE As Long = &H10
    Const BM_CLICK As Long = &HF5
    Const GW_CHILD As Long = 5
    Const GW_HWNDNEXT As Long = 2
    Const GWL_STYLE As Long = -16
    Const BS_DEFPUSHBUTTON As Long = &H1
    On Error Resume Next
    Dim myPid As Long
    myPid = GetCurrentProcessId()
    Dim fg As Long
    fg = GetForegroundWindow()
    Dim hDlg As Long
    hDlg = FindWindowExW(0, 0, StrPtr("#32770"), 0)
    Do While hDlg <> 0
        Dim pid As Long
        pid = 0
        GetWindowThreadProcessId hDlg, pid
        If pid = myPid And fg <> 0 And hDlg = fg Then
            ' prefer the default button so the dialog runs its own OK path
            Dim hChild As Long
            Dim btn As Long
            btn = 0
            hChild = GetWindow(hDlg, GW_CHILD)
            Do While hChild <> 0
                If (GetWindowLongW(hChild, GWL_STYLE) And BS_DEFPUSHBUTTON) <> 0 Then
                    btn = hChild
                    Exit Do
                End If
                hChild = GetWindow(hChild, GW_HWNDNEXT)
            Loop
            If btn <> 0 Then
                PostMessageW btn, BM_CLICK, 0, 0
            Else
                PostMessageW hDlg, WM_CLOSE, 0, 0
            End If
        End If
        hDlg = FindWindowExW(0, hDlg, StrPtr("#32770"), 0)
    Loop
    ' last resort only (localized caption); not the primary path any more
    Dim dlgName As String
    dlgName = ChrW$(&H65B0) & ChrW$(&H5EFA) & ChrW$(&H5DE5) & ChrW$(&H7A0B)
    Dim hDlg2 As Long
    hDlg2 = FindWindowW(0, StrPtr(dlgName))
    If hDlg2 <> 0 Then PostMessageW hDlg2, WM_CLOSE, 0, 0
End Sub

Private Function ToolRemoveProject(ByVal name As String) As String
    ' ---------------------------------------------------------------------------
    ' TASK-P1-090 (receipt honesty): this function used to report ok:false /
    ' unexpected_error on a SUCCESSFUL removal. Two independent defects:
    '   (1) it read p.Name AFTER m_vbe.VBProjects.Remove p. Remove disconnects the
    '       VBProject object, so that read raises on EVERY successful removal, and the
    '       Fail handler wrapped the raise into ok:false/unexpected_error even though
    '       the project really was gone (measured: P1-079, state axis showed no project);
    '   (2) a legal "nothing to remove" state (no project open, or a name that does
    '       not exist) also came out as unexpected_error instead of a structured refusal.
    ' The identity is now read BEFORE the removal, success is only claimed when the
    ' open-project count really dropped, and the legal states are structured refusals.
    ' ---------------------------------------------------------------------------
    If m_vbe Is Nothing Then
        ToolRemoveProject = RemoveProjectRefusal("ide_not_connected", _
            "the add-in has no live VBIDE connection", name)
        Exit Function
    End If
    If m_vbe.VBProjects.Count = 0 Then
        ToolRemoveProject = RemoveProjectRefusal("no_project", _
            "there is no project to remove: the IDE is already in the no-project state", name)
        Exit Function
    End If
    Dim p As VBIDE.VBProject
    Dim q As VBIDE.VBProject
    If Len(name) = 0 Then
        Set p = m_vbe.ActiveVBProject
    Else
        ' deliberately NOT FindProject(): its Err.Raise("Project not found") would once
        ' again surface as unexpected_error instead of a structured refusal
        For Each q In m_vbe.VBProjects
            If LCase$(q.Name) = LCase$(name) Then
                Set p = q
                Exit For
            End If
        Next
    End If
    If p Is Nothing Then
        If Len(name) = 0 Then
            ToolRemoveProject = RemoveProjectRefusal("no_active_project", _
                "the IDE has open projects but none of them is the active project", name)
        Else
            ToolRemoveProject = RemoveProjectRefusal("project_not_found", _
                "no open project is named '" & name & "'", name)
        End If
        Exit Function
    End If
    ' capture the identity BEFORE the removal (defect 1 above)
    Dim removedName As String
    Dim removedFile As String
    removedName = p.Name
    On Error Resume Next
    removedFile = p.FileName
    On Error GoTo 0
    Dim beforeCount As Long
    beforeCount = m_vbe.VBProjects.Count
    m_vbe.VBProjects.Remove p
    Dim afterCount As Long
    afterCount = m_vbe.VBProjects.Count
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    ' (3) never fake a success: the project list must really have shrunk.
    If afterCount >= beforeCount Then
        d("ok") = False
        d("reason") = "remove_project_not_effective"
        d("tool") = "remove_project"
        d("detail") = "VBProjects.Remove returned no error but the open-project count did not drop (" & CStr(beforeCount) & " -> " & CStr(afterCount) & ")"
        d("note") = "no success is claimed: the project may still be open"
        ToolRemoveProject = JsonStringify(d)
        Exit Function
    End If
    Set d("remainingProjects") = RemainingProjectNames()
    d("ok") = True
    d("removed") = removedName
    If Len(removedFile) > 0 Then d("removedFile") = removedFile
    d("closed") = True
    d("projectCountBefore") = beforeCount
    d("projectCountAfter") = afterCount
    d("note") = "the project was removed from the IDE's project group"
    ToolRemoveProject = JsonStringify(d)
End Function

' TASK-P1-090: a legal "cannot remove" condition (no project open / no such name) is a
' STRUCTURED refusal, never ok:false/unexpected_error. Nothing was removed when this
' returns, so the note says exactly that.
Private Function RemoveProjectRefusal(ByVal reason As String, ByVal detail As String, ByVal name As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("tool") = "remove_project"
    If Len(name) > 0 Then d("requestedName") = name
    d("detail") = detail
    d("note") = "nothing was removed: the IDE project group is unchanged"
    RemoveProjectRefusal = JsonStringify(d)
End Function

' TASK-P1-090: the projects still open AFTER a removal. Read from the live collection (and
' never from the removed object, which is disconnected).
Private Function RemainingProjectNames() As Object
    Dim c As Object
    Set c = New Collection
    Dim p As VBIDE.VBProject
    For Each p In m_vbe.VBProjects
        c.Add p.Name
    Next
    Set RemainingProjectNames = c
End Function

Private Function ProjectListArr() As Object
    Dim c As Object
    Set c = New Collection
    Dim p As VBIDE.VBProject
    For Each p In m_vbe.VBProjects
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = p.Name
        d("fileName") = p.FileName
        d("type") = ProjectTypeName(p)
        d("isStartProject") = False
        On Error Resume Next
        If m_vbe.VBProjects.StartProject Is p Then d("isStartProject") = True
        On Error GoTo 0
        c.Add d
    Next
    Set ProjectListArr = c
End Function

Private Function FindProject(ByVal name As String) As VBIDE.VBProject
    Dim p As VBIDE.VBProject
    For Each p In m_vbe.VBProjects
        If LCase$(p.Name) = LCase$(name) Then
            Set FindProject = p
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 46, "CodeTools", "Project not found: " & name
End Function

' =====================================================================
' component tools
' =====================================================================

Private Function ToolGetComponentList() As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim proj As VBIDE.VBProject
    Set proj = GetProject()
    root("project") = proj.Name
    Dim arr As Object
    Set arr = New Collection
    Dim p As VBIDE.VBProject
    Set p = proj
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = comp.Name
        d("type") = ComponentTypeName(comp.Type)
        On Error Resume Next
        Err.Clear
        d("file") = comp.FileNames(1)
        d("isDirty") = comp.IsDirty
        d("fileCount") = comp.FileCount
        d("designerID") = comp.DesignerID
        d("hasOpenDesigner") = comp.HasOpenDesigner
        d("description") = comp.Description
        d("helpContextID") = comp.HelpContextID
        d("iconState") = comp.IconState
        On Error GoTo 0
        Dim cm As VBIDE.CodeModule
        Set cm = comp.CodeModule
        If Not cm Is Nothing Then
            d("lines") = cm.CountOfLines
            d("declarationLines") = cm.CountOfDeclarationLines
            Set d("procedures") = ProcsOf(comp)
        End If
        arr.Add d
    Next
    Set root("components") = arr
    ToolGetComponentList = JsonStringify(root)
End Function

' Card TASK-P1-011: the leading header the code module does not expose
' (Attribute VB_Name = "...", VERSION ...). Returns "" when the pane cannot be
' located inside the file, since guessing a header is worse than omitting it.
Private Function ComponentHeaderOf(ByVal name As String, ByVal paneText As String) As String
    On Error Resume Next
    ComponentHeaderOf = ""
    Dim path As String
    path = DiskPathOf(name, "")
    If Len(path) = 0 Then Exit Function
    If Dir$(path) = "" Then Exit Function
    Dim disk As String
    disk = Replace(Build.ReadLogAnsi(path), vbCrLf, vbLf)
    Dim plines() As String
    plines = Split(Replace(paneText, vbCrLf, vbLf), vbLf)
    If UBound(plines) < 0 Then Exit Function
    Dim first As String
    first = plines(0)
    Dim dlines() As String
    dlines = Split(disk, vbLf)
    Dim i As Long
    Dim hdr As String
    hdr = ""
    For i = 0 To UBound(dlines)
        MaybePump
        If dlines(i) = first Then Exit For
        hdr = hdr & dlines(i) & vbCrLf
    Next
    If i > UBound(dlines) Then Exit Function
    ComponentHeaderOf = hdr
End Function

Private Function ToolGetComponentCode(ByVal name As String, ByVal includeLineNumbers As Boolean, _
                                      ByVal coordSpace As String, ByVal procName As String, ByVal symbolName As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 21, "CodeTools", "Component has no code module: " & name
    ' TASK-P1-024 gamma-1/gamma-2: coords picks the line space of the returned text
    ' (B+C batch default is "ide": pure pane text, IDE line numbers). "disk" keeps the
    ' A-batch behaviour byte-for-byte (header prepended, disk numbering) so an old
    ' caller only has to pass coords:"disk". proc:/symbol: narrows the range.
    Dim cs As String
    cs = LCase$(Trim$(coordSpace))
    If cs <> "disk" Then cs = "ide"
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("ideLineCount") = cm.CountOfLines
    root("coordSpace") = cs
    Dim wholePane As String
    ' TASK-P1-062: an empty component has CountOfLines = 0, and Lines(1, 0) raises VB
    ' error 5 ("invalid procedure call"). Reading nothing must return "" instead.
    If cm.CountOfLines > 0 Then wholePane = cm.Lines(1, cm.CountOfLines)
    Dim sIde As Long, eIde As Long, sliced As Boolean
    sIde = 1
    eIde = cm.CountOfLines
    sliced = False
    If Len(Trim$(procName)) > 0 Then
        If Not ProcRangeOf(cm, Trim$(procName), sIde, eIde) Then
            Set root("coords") = CoordsDictOfName(comp.Name, 1, cs)
            root("ok") = False
            root("reason") = "proc_not_found"
            root("proc") = procName
            root("lineSpace") = cs
            root("note") = "no procedure named '" & procName & "' in " & comp.Name & "; nothing was read"
            ToolGetComponentCode = JsonStringify(root)
            Exit Function
        End If
        sliced = True
    ElseIf Len(Trim$(symbolName)) > 0 Then
        If Not SymbolRangeOf(cm, Trim$(symbolName), sIde, eIde) Then
            Set root("coords") = CoordsDictOfName(comp.Name, 1, cs)
            root("ok") = False
            root("reason") = "symbol_not_found"
            root("symbol") = symbolName
            root("lineSpace") = cs
            root("note") = "no identifier named '" & symbolName & "' in " & comp.Name & "; nothing was read"
            ToolGetComponentCode = JsonStringify(root)
            Exit Function
        End If
        sliced = True
    End If
    ' TASK-P1-020 alpha (coords honesty): CodeModule.Lines() omits the file's
    ' Attribute/VERSION header, so this tool PREPENDS it (TASK-P1-011); the emitted
    ' line numbers are therefore DISK lines. coords says exactly that (default =
    ' "disk"), and headerOffset = diskLine - ideLine holds for the first line.
    Dim cd As Object
    Set cd = CoordsDictOfName(comp.Name, sIde, cs)
    Set root("coords") = cd
    ' TASK-P1-161 (R4): VB6 requires a LOCAL variable to be declared BEFORE its first
    ' use, so the accumulated map is resolved here - above every use of it.
    Dim em As Object
    Set em = PaneDiskMapCached(name)
    ' TASK-P1-020 rework: say which space the numbered output is in AND which space the
    ' edit tools take, so a caller cannot lose track (an independent verifier read the
    ' numbered lines as edit lines and reported a false defect).
    root("lineNumbersAre") = cs
    root("editToolsTake") = "ide"
    root("convertToEditLine") = "ideLine = diskLine - coords.headerOffset"
    ' TASK-P1-161 (R4): the mapping basis belongs at the TOP level of the one tool whose
    ' output is NUMBERED in disk space, and the hidden lines are named so the gaps in the
    ' numbering can be reproduced instead of being taken on trust.
    If cs = "disk" Then
        root("mappingBasis") = em("basis")
        root("mappingResolved") = em("resolved")
        root("hiddenInlineAttributeLines") = em("hiddenCount")
        Set root("hiddenInlineAttributeDiskLines") = em("hiddenDiskLines")
        If Not CBool(em("resolved")) Then Set root("mappingDeviationPaneLines") = em("deviationPaneLines")
        root("diskLineNote") = "in the DISK spaces the emitted line numbers skip the hidden inline Attribute lines, so they are not always consecutive; hiddenInlineAttributeDiskLines lists the disk lines the code pane does not show."
    End If
    If sliced Then
        root("startLine") = sIde
        root("endLine") = eIde
        If Len(Trim$(procName)) > 0 Then
            root("proc") = procName
        Else
            root("symbol") = symbolName
        End If
    End If
    Dim off As Long
    off = 0
    On Error Resume Next
    off = CLng(cd("headerOffset"))
    On Error GoTo 0
    ' TASK-P1-161 (R4): the DISK line of every EMITTED line comes from the accumulated
    ' pane->disk map, because a hidden inline Attribute line inside the module shifts
    ' everything below it. hl is the number of header lines this tool prepends.
    Dim headerText As String
    Dim hl As Long
    headerText = ""
    hl = 0
    Dim code As String
    Dim baseNo As Long
    If cs = "disk" And Not sliced Then
        headerText = ComponentHeaderOf(name, wholePane)
        hl = HeaderLineCountOf(headerText)
        code = headerText & wholePane
        baseNo = 1
    Else
        code = SliceLines(wholePane, sIde, eIde)
        If cs = "disk" Then
            baseNo = sIde + off
        Else
            baseNo = sIde
        End If
    End If
    root("includeLineNumbers") = includeLineNumbers
    root("lineCount") = LineCount(code)
    If cm.CountOfLines = 0 Then
        root("lineCount") = 0
        root("text") = ""
        root("textSha256") = Sha256.Sha256HexTextUtf8("")
        root("sha256") = root("textSha256")
        ToolGetComponentCode = JsonStringify(root)
        Exit Function
    End If
    If Not includeLineNumbers Then
        root("text") = code
        root("textSha256") = Sha256.Sha256HexTextUtf8(code)
        root("sha256") = root("textSha256")
        ToolGetComponentCode = JsonStringify(root)
        Exit Function
    End If
    Dim lines() As String
    lines = Split(code, vbCrLf)
    Dim out As String
    Dim i As Long
    ' TASK-P1-048: same quadratic accumulation as SliceLines (one append per source
    ' line). Build the numbered lines into an array and Join ONCE (reference
    ' technique: array + Join). Produced text is byte-identical.
    Dim ub As Long
    ub = -1
    On Error Resume Next
    ub = UBound(lines)
    On Error GoTo 0
    out = ""
    If ub >= 0 Then
        Dim parts() As String
        Dim k As Long
        ReDim parts(0 To ub)
        k = 0
        Dim thisNo As Long
        For i = 0 To ub
            MaybePump
            If i = ub Then
                If Len(lines(i)) = 0 Then Exit For
            End If
            ' TASK-P1-161 (R4): number each emitted line from the ACCUMULATED map. In the
            ' whole-module disk space the prepended header occupies disk lines 1..hl and
            ' the pane lines follow; in a slice the pane line is sIde + i.
            thisNo = baseNo + i
            If cs = "disk" Then
                If sliced Then
                    If ExactDiskLineOf(em, sIde + i, 0) > 0 Then thisNo = ExactDiskLineOf(em, sIde + i, 0)
                Else
                    If i < hl Then
                        thisNo = i + 1
                    Else
                        If ExactDiskLineOf(em, i - hl + 1, 0) > 0 Then thisNo = ExactDiskLineOf(em, i - hl + 1, 0)
                    End If
                End If
            End If
            parts(k) = Right$("    " & CStr(thisNo), 5) & "| " & lines(i) & vbCrLf
            k = k + 1
        Next
        If k > 0 Then
            If k <= ub Then ReDim Preserve parts(0 To k - 1)
            out = Join(parts, "")
        End If
    End If
    root("text") = out
    root("textSha256") = Sha256.Sha256HexTextUtf8(out)
    root("sha256") = root("textSha256")
    ToolGetComponentCode = JsonStringify(root)
End Function

' The number of lines in a header text (it ends with a CRLF, so the trailing empty
' element is not a line).
Private Function HeaderLineCountOf(ByVal headerText As String) As Long
    HeaderLineCountOf = 0
    If Len(headerText) = 0 Then Exit Function
    Dim a() As String
    a = Split(Replace(Replace(headerText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    If Len(a(UBound(a))) = 0 Then
        HeaderLineCountOf = UBound(a)
    Else
        HeaderLineCountOf = UBound(a) + 1
    End If
End Function


' TASK-P1-024 gamma-2 helpers -------------------------------------------------
Private Function SliceLines(ByVal pane As String, ByVal sLine As Long, ByVal eLine As Long) As String
    ' TASK-P1-048: this used to accumulate with "out = out & lines(i) & vbCrLf" once
    ' per line. That is quadratic and is the real ~12 s of the 1 MB whole-module
    ' get_component_code (28,814 appends, each copying the whole prefix); P4-047
    ' attributed that cost to the JSON step, but the JSON escaping itself measures
    ' ~130 ms/MB. Collect the wanted lines into an array and Join ONCE - the
    ' user-verified reference technique (a JSON module in the reference VB6 application: build a
    ' piece array, Join once, never "s = s & ..." in a loop). Output is identical:
    ' the old loop appended a CRLF after every line and then dropped the last one,
    ' which is exactly Join(parts, vbCrLf).
    Dim lines() As String, i As Long, k As Long
    lines = Split(pane, vbCrLf)
    If sLine < 1 Then sLine = 1
    If eLine > UBound(lines) + 1 Then eLine = UBound(lines) + 1
    If eLine < sLine Then
        SliceLines = ""
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To eLine - sLine)
    k = 0
    For i = sLine To eLine
        MaybePump
        parts(k) = lines(i - 1)
        k = k + 1
    Next
    SliceLines = Join(parts, vbCrLf)
End Function

Private Function ProcRangeOf(ByVal cm As VBIDE.CodeModule, ByVal procName As String, ByRef sLine As Long, ByRef eLine As Long) As Boolean
    On Error Resume Next
    ProcRangeOf = False
    Dim kinds(0 To 3) As Long
    kinds(0) = vbext_pk_Proc
    kinds(1) = vbext_pk_Get
    kinds(2) = vbext_pk_Let
    kinds(3) = vbext_pk_Set
    Dim i As Long, st As Long, n As Long
    For i = 0 To 3
        Err.Clear
        st = 0
        st = cm.ProcStartLine(procName, kinds(i))
        If Err.Number = 0 Then
            If st > 0 Then
                Err.Clear
                n = cm.ProcCountLines(procName, kinds(i))
                If Err.Number = 0 Then
                    If n > 0 Then
                        sLine = st
                        eLine = st + n - 1
                        ProcRangeOf = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function SymbolRangeOf(ByVal cm As VBIDE.CodeModule, ByVal sym As String, ByRef sLine As Long, ByRef eLine As Long) As Boolean
    On Error Resume Next
    SymbolRangeOf = False
    Dim pane As String, lines() As String, i As Long, hit As Long
    ' TASK-P1-062: guard the empty-component case (CountOfLines = 0 => Lines(1,0) = error 5)
    If cm.CountOfLines > 0 Then pane = cm.Lines(1, cm.CountOfLines)
    lines = Split(pane, vbCrLf)
    hit = 0
    For i = 0 To UBound(lines)
        MaybePump
        If HasWholeWord(lines(i), sym) Then
            hit = i + 1
            Exit For
        End If
    Next
    If hit = 0 Then Exit Function
    Dim p As String
    p = cm.ProcOfLine(hit, vbext_pk_Proc)
    If Len(p) > 0 Then
        If ProcRangeOf(cm, p, sLine, eLine) Then
            SymbolRangeOf = True
            Exit Function
        End If
    End If
    sLine = hit
    eLine = hit
    SymbolRangeOf = True
End Function

Private Function HasWholeWord(ByVal line As String, ByVal word As String) As Boolean
    Dim lc As String, wc As String, posn As Long, before As String, after As String
    HasWholeWord = False
    wc = LCase$(word)
    If Len(wc) = 0 Then Exit Function
    lc = LCase$(line)
    posn = InStr(1, lc, wc, vbBinaryCompare)
    Do While posn > 0
        before = ""
        after = ""
        If posn > 1 Then before = Mid$(lc, posn - 1, 1)
        If posn + Len(wc) <= Len(lc) Then after = Mid$(lc, posn + Len(wc), 1)
        If Not IsWordChar(before) And Not IsWordChar(after) Then
            HasWholeWord = True
            Exit Function
        End If
        posn = InStr(posn + 1, lc, wc, vbBinaryCompare)
    Loop
End Function

Private Function IsWordChar(ByVal c As String) As Boolean
    IsWordChar = (Len(c) > 0 And InStr("abcdefghijklmnopqrstuvwxyz0123456789_", c) > 0)
End Function

Private Function ToolReplaceCode(ByVal name As String, ByVal operation As String, ByVal startLine As Long, ByVal endLine As Long, _
                                 ByVal code As String, ByVal expectedSha As String, ByVal validateMode As String, _
                                 ByVal snapshotOn As Boolean) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 21, "CodeTools", "Component has no code module: " & name
    ' TASK-P1-049: totalMs must cover the WHOLE call. It used to start just before
    ' the write dispatch, so a slow read/predict/lint still reported totalMs=0
    ' (P4-039). tEntry is the first statement of the tool; the marks below split
    ' the call into read / predict / lint.
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long
    tEntry = GetTickCount()
    Dim op As String
    op = LCase$(operation)
    ' TASK-P1-126 (MCP-07): a LINE-NUMBER edit must pin the text it was measured
    ' against. The user lost code once by replacing a remembered range after the file
    ' had shifted ("stale line numbers hit another procedure"). The guard used to be
    ' OPTIONAL for every operation; for replace/delete it is now REQUIRED, and the
    ' refusal carries the sha the caller needs, so the fix is a single retry.
    ' insert/append/set keep the old optional behaviour: an insert only shifts text, it
    ' cannot destroy the code the caller measured (and set= replaces everything, where
    ' a whole-text sha carries no information).
    If (op = "replace" Or op = "delete") And Len(Trim$(expectedSha)) = 0 Then
        ToolReplaceCode = MissingShaRefusal(comp.Name, op, startLine, endLine, cm)
        Exit Function
    End If
    Dim norm As String
    norm = NormalizeNewlines(code)
    ' TASK-P1-126 (MCP-06/MCP-04): the PRE-WRITE structural gate. A block carrying its
    ' own header plus a redundant "End Function" used to be WRITTEN (only the linter
    ' complained, VB6-E105) and cost two failed builds.
    Dim pairRef As String
    pairRef = PreWritePairingRefusal("replace_code", comp.Name, norm, "code (operation=" & op & ")")
    If Len(pairRef) > 0 Then
        ToolReplaceCode = pairRef
        Exit Function
    End If
    Dim linesArr() As String
    linesArr = Split(norm, vbLf)
    Dim nLines As Long
    nLines = UBound(linesArr) + 1
    If nLines > 0 Then
        If linesArr(nLines - 1) = "" Then nLines = nLines - 1
    End If
    Dim reqStart As Long, reqEnd As Long
    reqStart = startLine
    reqEnd = endLine
    Dim countOfLines As Long
    countOfLines = cm.CountOfLines
    ' ---- TASK-P1-020 beta-1: an out-of-range replace/delete is REFUSED, never
    ' clamped. The old path silently degraded to an insert and still answered
    ' deletedLines:0, which is how a wrong line number lost code (in a real project).
    If op = "replace" Or op = "delete" Then
        If reqStart > countOfLines Or reqEnd > countOfLines Then
            ToolReplaceCode = GuardReject(comp.Name, op, "range_beyond_end", _
                "start_line/end_line are past CountOfLines (" & CStr(countOfLines) & "); the IDE was not touched. Re-read the component with get_component_code and edit inside 1.." & CStr(countOfLines) & ", or use insert/append to add lines.", _
                reqStart, reqEnd, countOfLines)
            Exit Function
        End If
    End If
    ' ---- TASK-P1-020 beta-2: never cut a VB line continuation in two ----
    Dim curPane As String
    ' TASK-P1-062: guard the empty-component case (CountOfLines = 0 => Lines(1,0) = error 5)
    If cm.CountOfLines > 0 Then curPane = cm.Lines(1, cm.CountOfLines)
    ' TASK-P1-024 gamma-3: the same optimistic-concurrency guard the anchor tools use
    GuardSha curPane, expectedSha
    Dim curLines() As String
    curLines = TextLines(curPane)
    If op = "replace" Or op = "delete" Then
        If IsContinuationTail(curLines, reqStart) Then
            ToolReplaceCode = GuardReject(comp.Name, op, "inside_continuation", _
                "line " & CStr(reqStart) & " is the tail of a logical line: line " & CStr(reqStart - 1) & " ends with a VB continuation _. Replacing or deleting it in place would split that logical line. Widen the range to the whole logical statement, or edit by anchor on its first line.", _
                reqStart, reqEnd, countOfLines)
            Exit Function
        End If
    ElseIf op = "insert" Then
        If IsContinuationTail(curLines, reqStart) Then
            ToolReplaceCode = GuardReject(comp.Name, op, "inside_continuation", _
                "the insertion point is inside a logical line: line " & CStr(reqStart - 1) & " ends with a VB continuation _. Insert before the first line of that logical statement, or after its last line.", _
                reqStart, reqStart, countOfLines)
            Exit Function
        End If
    ElseIf op = "append" Then
        If IsContinuationTail(curLines, countOfLines + 1) Then
            ToolReplaceCode = GuardReject(comp.Name, op, "inside_continuation", _
                "the last line ends with a VB continuation _, so appending would extend that logical statement. Close the statement first.", _
                countOfLines, countOfLines, countOfLines)
            Exit Function
        End If
    End If
    ' ---- TASK-P1-020 beta-3: lint the RESULT before it is written (warn-only) ----
    tRead = GetTickCount()
    Dim adjStart As Long, adjEnd As Long
    adjStart = reqStart
    adjEnd = reqEnd
    If (op = "replace" Or op = "delete") Then
        If adjStart < 1 Then adjStart = 1
        If adjEnd < adjStart Then adjEnd = adjStart
    End If
    Dim predicted As String
    predicted = PredictedReplaceText(curPane, op, adjStart, adjEnd, linesArr, nLines)
    tPredict = GetTickCount()
    Dim pack As Object
    ' TASK-P1-051: above the region threshold, lint targets the edited procedure
    ' instead of the whole module (the receipt states the scope explicitly).
    Dim lintTarget As Long
    If countOfLines > LintRegionThreshold() Then lintTarget = LintRegionTarget(op, adjStart, countOfLines)
    Set pack = LintPackScoped(predicted, KindOfComponent(comp), lintTarget)
    tLint = GetTickCount()
    ' TASK-P1-024 gamma-4: validate:'strict' refuses BEFORE the IDE is touched
    Dim rcref As String
    rcref = StrictRefusal(comp.Name, "replace_code", pack, validateMode)
    If Len(rcref) > 0 Then
        ToolReplaceCode = rcref
        Exit Function
    End If
    Dim clamped As Boolean
    clamped = False
    Dim apStart As Long, apEnd As Long
    apStart = 1
    apEnd = 0
    Dim summary As Object
    Set summary = CreateObject("Scripting.Dictionary")
    summary("readMs") = tRead - tEntry
    summary("predictMs") = tPredict - tRead
    summary("lintMs") = tLint - tPredict
    summary("component") = comp.Name
    summary("operation") = op
    ' TASK-P1-075: m_lastWriteHunks is module state; reset it per call so a branch that
    ' never writes through SetCompTextMinimal can not report the PREVIOUS call's hunk count.
    ' TASK-P1-078: reset the whole write-phase block here, not just the hunk count. The raw
    ' insert/append/delete branches below never go through SetCompTextMinimal and never
    ' publish the phases, so without these lines they would leave the PREVIOUS call's
    ' diffMs/joinMs/delMs/insMs/maxBlockLines in module state for any later reader
    ' (exactly the P1-075 defect class, one field further out).
    ' tRaw/wcRaw time and count the raw single-block operations (insert/append/delete).
    m_lastWriteHunks = 0
    m_lastDiffMs = 0
    m_lastJoinMs = 0
    m_lastDelMs = 0
    m_lastInsMs = 0
    m_lastMaxBlock = 0
    Dim tRaw As Long, wcRaw As Long
    Dim i As Long
    ' TASK-P1-087: opt-in pre-edit snapshot, taken exactly ONCE before any write. The four
    ' RAW branches below (replace / insert / append / delete) - and set's degenerate branch -
    ' write through CodeModule directly and never reached the snapshot funnel, which is why
    ' undo_edit could not undo a plain replace_code replace. Default off => no new file and a
    ' byte-identical receipt.
    If snapshotOn Then UndoSnapshot comp.Name, curPane
    Select Case op
        Case "set"
            Dim oldCount As Long
            oldCount = cm.CountOfLines
            ' TASK-P1-059(c): measured on a 27,000-line / 1.2 MB PURE-ASCII module, the
            ' WriteMs of DeleteLines+InsertLines is a SUPER-LINEAR function of the BLOCK
            ' SIZE and is INDEPENDENT of the block POSITION: 1 line = 0 ms at k=8/7000/
            ' 13500/20000/26990; 50 lines = 15 ms; 500 lines = 31 ms; 5,000 lines = 844 ms;
            ' the whole 27,000-line module = 21,593 ms. The old set branch always paid the
            ' last number (measured 19-27 s of set write time) no matter where the caller's
            ' edit was. Writing only the changed range (common prefix/suffix, <= 1 delete +
            ' <= 1 insert) produces the IDENTICAL pane text for at most one line of cost.
            Dim tWriteS As Long
            Dim setWrites As Long
            ' TASK-P1-070: open the phase window for THIS write. The degenerate
            ' branch below does not go through SetCompTextMinimal, so it resets
            ' the phases itself instead of inheriting the previous call's.
            m_lastDiffMs = 0
            m_lastJoinMs = 0
            m_lastDelMs = 0
            m_lastInsMs = 0
            m_lastMaxBlock = 0
            tWriteS = GetTickCount()
            If oldCount > 0 And nLines > 0 Then
                Dim curSet As String
                curSet = CompText(comp.Name)
                setWrites = SetCompTextMinimal(comp.Name, curSet, JoinLines(linesArr, nLines), snapshotOn)
            Else
                ' TASK-P1-059(c) guard: the two degenerate shapes are NOT representable as a
                ' minimal range write - the common-prefix helper would issue DeleteLines on
                ' a 0-line module (VB error) or insert a phantom blank line when the
                ' requested text is empty. Those keep the historical two calls verbatim.
                ' TASK-P1-075: one logical block (<=1 delete + <=1 insert) => hunks = 1.
                m_lastWriteHunks = 1
                setWrites = 0
                Dim tPhs As Long
                tPhs = GetTickCount()
                If oldCount > 0 Then
                    tPhs = GetTickCount()
                    cm.DeleteLines 1, oldCount
                    m_lastDelMs = m_lastDelMs + (GetTickCount() - tPhs)
                    If oldCount > m_lastMaxBlock Then m_lastMaxBlock = oldCount
                    setWrites = setWrites + 1
                End If
                If nLines > 0 Then
                    tPhs = GetTickCount()
                    Dim wholeText As String
                    wholeText = JoinLines(linesArr, nLines)
                    m_lastJoinMs = m_lastJoinMs + (GetTickCount() - tPhs)
                    tPhs = GetTickCount()
                    cm.InsertLines 1, wholeText
                    m_lastInsMs = m_lastInsMs + (GetTickCount() - tPhs)
                    If nLines > m_lastMaxBlock Then m_lastMaxBlock = nLines
                    setWrites = setWrites + 1
                End If
            End If
            summary("writeMs") = GetTickCount() - tWriteS
            AddWritePhases summary
            summary("writeCalls") = setWrites
            summary("hunks") = m_lastWriteHunks
            summary("insertedLines") = nLines
            summary("deletedLines") = oldCount
            apStart = 1
            apEnd = nLines
        Case "replace"
            startLine = adjStart
            endLine = adjEnd
            Dim rangeCount As Long
            rangeCount = endLine - startLine + 1
            ' TASK-P1-043: ONE range replace instead of a per-line loop.
            ' VB6 fully REPARSES/REFORMATS the whole module after EVERY ReplaceLine /
            ' InsertLines, so the old loop paid one whole-module reformat per line: a
            ' 5-line edit on a 28,814-line (1 MB) module issued 5 ReplaceLine + 1
            ' InsertLines = 6 reformats, measured by the user as about 60 s. Now: at most
            ' one DeleteLines plus at most one InsertLines, i.e. writeCalls <= 2.
            ' Semantics of replacedLines/deletedLines/insertedLines and the clamp are
            ' UNCHANGED (see the summary block below).
            ' NOTE the name wipeCount, not delCount: VB6 has no block scope, and the
            ' "delete" case below already Dims delCount in this same procedure - reusing
            ' that name is a compile error (duplicate declaration in the current scope).
            Dim wipeCount As Long, writeCalls As Long
            wipeCount = rangeCount
            If startLine + wipeCount - 1 > cm.CountOfLines Then wipeCount = cm.CountOfLines - startLine + 1
            If wipeCount < 0 Then wipeCount = 0
            Dim tWrite0 As Long
            ' TASK-P1-070: this branch writes through the CodeModule directly, so
            ' measure its phases here (same field names as the minimal writer).
            m_lastDiffMs = 0
            m_lastJoinMs = 0
            m_lastDelMs = 0
            m_lastInsMs = 0
            m_lastMaxBlock = 0
            m_lastWriteHunks = 1
            Dim tPh0 As Long
            tWrite0 = GetTickCount()
            If wipeCount > 0 Then
                tPh0 = GetTickCount()
                cm.DeleteLines startLine, wipeCount
                m_lastDelMs = m_lastDelMs + (GetTickCount() - tPh0)
                If wipeCount > m_lastMaxBlock Then m_lastMaxBlock = wipeCount
                writeCalls = writeCalls + 1
            End If
            If nLines > 0 Then
                tPh0 = GetTickCount()
                Dim repText As String
                repText = JoinLines(linesArr, nLines)
                m_lastJoinMs = m_lastJoinMs + (GetTickCount() - tPh0)
                tPh0 = GetTickCount()
                cm.InsertLines startLine, repText
                m_lastInsMs = m_lastInsMs + (GetTickCount() - tPh0)
                If nLines > m_lastMaxBlock Then m_lastMaxBlock = nLines
                writeCalls = writeCalls + 1
            End If
            summary("writeMs") = GetTickCount() - tWrite0
            AddWritePhases summary
            summary("writeCalls") = writeCalls
            summary("hunks") = m_lastWriteHunks
            If nLines < rangeCount Then
                summary("replacedLines") = nLines
                summary("deletedLines") = rangeCount - nLines
                summary("insertedLines") = 0
            Else
                summary("replacedLines") = rangeCount
                summary("deletedLines") = 0
                summary("insertedLines") = nLines - rangeCount
            End If
            apStart = startLine
            apEnd = startLine + nLines - 1
        Case "insert"
            If startLine < 1 Then startLine = 1
            If startLine > cm.CountOfLines + 1 Then
                clamped = True
                startLine = cm.CountOfLines + 1
            End If
            ' TASK-P1-075: insert/append/delete are raw single-block writes too - they now
            ' report the SAME write-phase fields (writeMs/writeCalls/hunks) as set/replace,
            ' so every minimal-write receipt carries hunks (CONTRACTS §3). hunks = 1.
            tRaw = GetTickCount()
            wcRaw = 0
            If nLines > 0 Then
                cm.InsertLines startLine, JoinLines(linesArr, nLines)
                wcRaw = 1
            End If
            m_lastWriteHunks = 1
            summary("writeMs") = GetTickCount() - tRaw
            summary("writeCalls") = wcRaw
            summary("hunks") = m_lastWriteHunks
            summary("insertedLines") = nLines
            apStart = startLine
            apEnd = startLine + nLines - 1
        Case "append"
            tRaw = GetTickCount()
            wcRaw = 0
            If nLines > 0 Then
                cm.InsertLines cm.CountOfLines + 1, JoinLines(linesArr, nLines)
                wcRaw = 1
            End If
            m_lastWriteHunks = 1
            summary("writeMs") = GetTickCount() - tRaw
            summary("writeCalls") = wcRaw
            summary("hunks") = m_lastWriteHunks
            summary("insertedLines") = nLines
            apStart = cm.CountOfLines - nLines + 1
            apEnd = cm.CountOfLines
        Case "delete"
            ' TASK-P1-020 beta-1: the range is already known to be inside the module
            ' (an out-of-range delete was refused above), so nothing is clamped here.
            startLine = adjStart
            endLine = adjEnd
            Dim delCount As Long
            delCount = endLine - startLine + 1
            tRaw = GetTickCount()
            cm.DeleteLines startLine, delCount
            m_lastWriteHunks = 1
            summary("writeMs") = GetTickCount() - tRaw
            summary("writeCalls") = 1
            summary("hunks") = m_lastWriteHunks
            summary("deletedLines") = delCount
            apStart = startLine
            apEnd = startLine - 1
        Case Else
            Err.Raise vbObjectError + 22, "CodeTools", "Invalid operation: " & operation & " (expected replace/insert/append/delete/set)"
    End Select
    If (op = "replace" Or op = "delete") Then
        If reqStart <> adjStart Or reqEnd <> adjEnd Then clamped = True
    End If
    ' ---- TASK-P1-020 alpha-3: the IDE-read count next to the text-derived one ----
    Dim paneAfter As String
    paneAfter = CompText(comp.Name)
    Dim textTotal As Long, ideTotal As Long
    textTotal = LineCount(paneAfter)
    ideTotal = cm.CountOfLines
    summary("newTotalLines") = ideTotal
    summary("ideTotalLines") = ideTotal
    summary("textTotalLines") = textTotal
    If ideTotal <> textTotal Then summary("syncWarning") = SyncWarningText(ideTotal, textTotal)
    ' ---- TASK-P1-020 alpha-2: requested vs applied range, clamp visible ----
    Set summary("requestedRange") = RangeDict(op, reqStart, reqEnd)
    Set summary("appliedRange") = AppliedRangeDict(apStart, apEnd)
    summary("clamped") = clamped
    Set summary("coords") = CoordsDictOfName(comp.Name, apStart, "ide")
    AddPaneDiskFields summary, comp.Name, paneAfter
    CopyLintPack summary, pack
    ' TASK-P1-049: totalMs covers the WHOLE call (entry -> receipt), so a slow call
    ' is self-explaining and can never report 0 ms again.
    ' TASK-P1-085: replace_code has its OWN receipt builder (it does not go through
    ' ChangeResult2), so the project-directory write manifest is added here explicitly -
    ' without this the _undo snapshot written by the text-write funnel would be invisible.
    Dim pdwRC As Object
    Set pdwRC = New Collection
    ManifestMergeUndo pdwRC
    Set summary("projectDirWrites") = pdwRC
    summary("totalMs") = GetTickCount() - tEntry
    ' TASK-P1-086 (F3): ReceiptJson serialises and adds serializeMs/receiptBytes; totalMs
    ' keeps its exact historical stamp point (just above) so old ledgers stay comparable.
    ToolReplaceCode = ReceiptJson(summary)
End Function

Private Function ToolAddComponent(ByVal compType As String, ByVal name As String, ByVal code As String) As String
    Dim ct As vbext_ComponentType
    Select Case LCase$(compType)
        Case "module"
            ct = vbext_ct_StdModule
        Case "class"
            ct = vbext_ct_ClassModule
        Case "form"
            ct = vbext_ct_MSForm
        Case "mdiform"
            ct = vbext_ct_VBMDIForm
        Case "resource"
            ct = vbext_ct_ResFile
        Case "usercontrol"
            ct = vbext_ct_UserControl
        Case "proppage"
            ct = vbext_ct_PropPage
        Case "userdoc"
            ct = vbext_ct_DocObject
        Case "designer"
            ct = vbext_ct_ActiveXDesigner
        Case Else
            Err.Raise vbObjectError + 47, "CodeTools", "Unknown component type: " & compType
    End Select
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim comp As VBIDE.VBComponent
    Dim createdFile As String
    createdFile = ""
    Dim createdVia As String
    createdVia = ""
    ' VBComponents.Add(vbext_ct_MSForm) raises E_FAIL (0x80004005) on installs
    ' where the Microsoft Form designer cannot be instantiated, while module,
    ' class, usercontrol and MDI form all add fine.
    ' TASK-P1-084: the recovery is now the IDE'S OWN Add Form command (the IDE
    ' builds its own default form); the synthesized .frm + AddFile remains only as
    ' the last-resort fallback inside AddControlFormFile.
    On Error Resume Next
    Set comp = p.VBComponents.Add(ct)
    Dim addErr As Long
    addErr = Err.Number
    Err.Clear
    On Error GoTo 0
    If addErr <> 0 Or comp Is Nothing Then
        If LCase$(compType) <> "form" Then
            Err.Raise vbObjectError + 46, "CodeTools", "VBComponents.Add failed for type " & compType & " (0x" & Hex$(addErr) & ")"
        End If
        Set comp = AddControlFormFile(p, name, createdFile, createdVia)
        If comp Is Nothing Then Err.Raise vbObjectError + 46, "CodeTools", "Could not create a form component"
    End If
    If Len(name) > 0 Then comp.Name = name
    If Len(code) > 0 Then
        If Not comp.CodeModule Is Nothing Then
            comp.CodeModule.AddFromString code
            ' VBComponents.Add seeds new modules with an Option
            ' Explicit header; de-duplicate when the supplied code
            ' brings its own (duplicate Option statements are
            ' compile errors that break make/run)
            Dim nLines As Long
            nLines = comp.CodeModule.CountOfLines
            If nLines > 1 Then
                Dim seenOE As Boolean
                seenOE = False
                Dim li As Long
                For li = 1 To nLines
                    MaybePump
                    If li > 12 Then Exit For
                    Dim t As String
                    t = LCase$(Trim$(comp.CodeModule.Lines(li, 1)))
                    If t = "option explicit" Then
                        If seenOE Then
                            comp.CodeModule.DeleteLines li, 1
                            Exit For
                        End If
                        seenOE = True
                    End If
                Next
            End If
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("created") = comp.Name
    SeNoteCreated comp.Name
    root("type") = ComponentTypeName(comp.Type)
    root("lines") = IIf(comp.CodeModule Is Nothing, 0, comp.CodeModule.CountOfLines)
    If Len(createdFile) > 0 Then root("file") = createdFile
    ' TASK-P1-084: how the form came into existence - "ide_dialog" (the IDE's own
    ' Add Form command built it: nothing written by us) or "synthesized" (last-resort
    ' fallback). Pure output: the inputSchema is unchanged.
    If Len(createdVia) > 0 Then
        root("createdVia") = createdVia
        If createdVia = "synthesized" Then root("ideRouteNote") = WinOps.WoConfirmStatus()
    End If
    ' TASK-P1-077: the write manifest is ALWAYS present (empty array when the call wrote no
    ' file). add_component can only write a file on the form fallback, and that fallback is
    ' refused unless explicitly opted in, so this is [] on every normal call.
    Dim writes As Object
    Set writes = New Collection
    ManifestAdd writes, createdFile
    ManifestMergeUndo writes
    Set root("projectDirWrites") = writes
    ToolAddComponent = JsonStringify(root)
End Function

' TASK-P1-084 (内部禁用路由说明 ruling 2 replacement): create the form
' through the IDE ITSELF. Two routes, IDE-first, no header synthesized by us:
'   1. the IDE's own Project > Add Form command (the IDE builds its own default
'      form; measured LIVE: 497 B / geometry 3015/120/465/4560 - our old
'      synthesized approximation was 3090/60/450/4680, i.e. a form the IDE would
'      never have produced);
'   2. last-resort fallback: the legacy synthesized .frm + AddFile below. It is
'      KEPT on purpose - removing it needs regression evidence and an S ruling -
'      and it stays gated by ProjectDirWritesKilled (ruling route 1 writes no file
'      at all, so that gate must not apply to it).
' outRoute reports which route ran: "ide_dialog" | "synthesized".
Private Function AddControlFormFile(ByVal proj As VBIDE.VBProject, ByVal name As String, ByRef outFile As String, ByRef outRoute As String) As VBIDE.VBComponent
    Dim compName As String
    compName = name
    If Len(compName) = 0 Then compName = NextFreeComponentName(proj, "Form")
    ' ---- route 1: the IDE creates the form (nothing written by us) ----------
    Dim ideComp As VBIDE.VBComponent
    Set ideComp = CreateFormViaIdeMenu(proj, compName, outRoute)
    If Not ideComp Is Nothing Then
        outFile = ""                    ' the IDE owns the file; we wrote nothing
        Set AddControlFormFile = ideComp
        Exit Function
    End If
    ' ---- fallback: legacy synthesized .frm (unchanged gate) -----------------
    outRoute = "synthesized"
    ' TASK-P1-077: this is the ONLY path that writes a file the caller did not name. Two facts
    ' pin down its gate: (1) it is reachable ONLY when the caller explicitly asked for a form
    ' (type="form") and VBComponents.Add(vbext_ct_MSForm) failed - measured LIVE on this machine
    ' (add_component {type:"form"} returns a "file"); a VB6 form IS a project member file, so
    ' there is no way to create one without writing <ProjectDir>\<name>.frm, and refusing it
    ' unconditionally would remove form creation altogether; (2) the .frm header / Begin..End
    ' layout block is synthesized by us - the route 内部禁用路由说明 ruling 2 vetoes.
    ' Gate therefore = the caller's explicit request is the opt-in, the receipt always names the
    ' file, the write is IMPOSSIBLE with no project directory (the old %TEMP% / "C:" fallback is
    ' gone: a project member file must live in the project) and the operator can kill it with
    ' VB6MCP_ALLOW_PROJECTDIR_WRITES=0 (ProjectDirWritesKilled).
    Dim dir As String
    dir = ProjectDir()
    If Len(dir) = 0 Then
        Err.Raise vbObjectError + 99, "CodeTools", "MCPREASON:project_dir_unavailable|" & _
            "save the project first (save_project with a path) and retry, or create the form in the IDE|" & _
            "the form fallback needs a project directory for the generated .frm and this project has no file on disk; nothing was written"
    End If
    If ProjectDirWritesKilled() Then
        Err.Raise vbObjectError + 101, "CodeTools", "MCPREASON:project_dir_write_not_allowed|" & _
            "unset the environment switch VB6MCP_ALLOW_PROJECTDIR_WRITES (it is 0/false/off in this VB6 process), or create the form in the IDE|" & _
            "the operator disabled project-directory writes, and creating a form through the AddFile fallback would write a generated .frm into the project directory; nothing was written"
    End If
    Dim full As String
    full = CtlJoinPath(dir, compName & ".frm")
    Dim n As Long
    n = 1
    Do While CtlFileExists(full)
        n = n + 1
        full = CtlJoinPath(dir, compName & CStr(n) & ".frm")
    Loop
    CtlWriteFile full, FormFileText(compName)
    Dim comp As VBIDE.VBComponent
    Set comp = proj.VBComponents.AddFile(full, False)
    outFile = full
    Set AddControlFormFile = comp
End Function

' TASK-P1-084: the IDE's own "Add Form" command (Project menu, localized
' "添加窗体" / English "Add Form"). It opens the IDE's template chooser, whose
' confirm button is "&Open" - a caption close_active_dialogs deliberately does NOT
' whitelist, so that tool cannot dismiss this box.
' ★The box is MODAL and runs its own message loop. A tool call that opens it is
' suspended INSIDE that loop: its DoEvents never returns, so the code after it can
' never click the button. The click therefore comes from the add-in's watchdog tick
' (WinOps.WoConfirmTick, driven by ServerForm.tmrWatch -> WinOps.WheelTick), which
' the modal loop DOES dispatch. Sequence: arm -> execute the menu command -> pump
' (the tick clicks; the modal closes; the pump returns) -> read the component the
' IDE created back out of the project.
' Failure is honest: nothing is clicked unless it is a visible #32770 of OUR process
' with an enabled confirm button, and no dialog we did not open is ever touched.
Private Function CreateFormViaIdeMenu(ByVal proj As VBIDE.VBProject, ByVal compName As String, ByRef route As String) As VBIDE.VBComponent
    On Error GoTo Fail
    Dim t0 As Long
    t0 = GetTickCount()
    Dim namesBefore As String
    namesBefore = ComponentNameSet(proj)
    ' Refuses when a dialog of ours is already up (never press a button in a box this
    ' call did not open) - the caller then falls back and reports why.
    If Not WinOps.WoConfirmArm(15000) Then Exit Function
    Dim ctrl As Object
    Set ctrl = FindBarControlDeep(FormAddMenuCands())
    If ctrl Is Nothing Then
        WinOps.WoConfirmDisarm
        Exit Function
    End If
    Dim why As String
    If ExecIfEnabledWhy(ctrl, "add_form", why) = 0 Then
        WinOps.WoConfirmDisarm
        Exit Function
    End If
    ' Pump until the watchdog's click lands (bounded by the arm deadline; the tick
    ' also dismisses the box at the deadline, so this cannot spin forever).
    Do While GetTickCount() - t0 < 20000
        DoEvents
        MaybeHonorCancel
        If WinOps.WoConfirmClicked() Then Exit Do
        If Not WinOps.WoConfirmArmed() Then Exit Do
        Sleep 40
    Loop
    WinOps.WoConfirmDisarm
    ' The IDE creates and names the component itself: find it by name difference.
    Dim i As Long, nm As String
    For i = 1 To 30
        nm = NewComponentName(proj, namesBefore)
        If Len(nm) > 0 Then Exit For
        DoEvents
        Sleep 60
    Next i
    If Len(nm) = 0 Then Exit Function
    Dim c As VBIDE.VBComponent
    Set c = proj.VBComponents(nm)
    If c Is Nothing Then Exit Function
    If Len(compName) > 0 Then
        On Error Resume Next
        c.Name = compName
        Err.Clear
        On Error GoTo Fail
        ' The IDE builds the form as "Form1"; renaming the COMPONENT does not rewrite
        ' the two designer properties that default to the original name (measured LIVE:
        ' after c.Name = compName the .frm still said Caption="Form1" and
        ' LinkTopic="Form1" - the very defect this card fixes in the fallback). Set
        ' both through the IDE's OWN property model - exactly what a user does in the
        ' Properties window (measured: both accept the new value and the IDE writes it
        ' to the .frm) - so nothing about the header is synthesized by us.
        AlignNewFormName c, compName
    End If
    route = "ide_dialog"
    Set CreateFormViaIdeMenu = c
    Exit Function
Fail:
    WinOps.WoConfirmDisarm
    Err.Clear
End Function

' Caption and LinkTopic both default to the name the form had when the IDE created
' it, and VB6 does NOT update them when the component is renamed. Best effort: a
' property the designer does not expose is left alone (never fatal).
Private Sub AlignNewFormName(ByVal c As VBIDE.VBComponent, ByVal compName As String)
    On Error Resume Next
    Dim props As Object
    Set props = c.Properties
    If props Is Nothing Then Exit Sub
    SetFormDesignerProp props, "Caption", compName
    SetFormDesignerProp props, "LinkTopic", compName
    Err.Clear
End Sub

Private Sub SetFormDesignerProp(ByVal props As Object, ByVal propName As String, ByVal v As String)
    On Error Resume Next
    props(propName).Value = v
    Err.Clear
End Sub

' "添加窗体" / "Add Form" - caption candidates, because the menu bar and the item
' are localized and VB6 Ids are not unique enough to name a control on their own.
Private Function FormAddMenuCands() As String
    FormAddMenuCands = ChrW$(&H6DFB) & ChrW$(&H52A0) & ChrW$(&H7A97) & ChrW$(&H4F53) & "|Add Form"
End Function

' "|a|b|" membership set of the project's component names.
Private Function ComponentNameSet(ByVal proj As VBIDE.VBProject) As String
    On Error Resume Next
    Dim s As String, c As VBIDE.VBComponent
    s = "|"
    For Each c In proj.VBComponents
        s = s & c.Name & "|"
    Next
    ComponentNameSet = s
End Function

' The first component whose name is NOT in beforeSet (the IDE named it).
Private Function NewComponentName(ByVal proj As VBIDE.VBProject, ByVal beforeSet As String) As String
    On Error Resume Next
    Dim c As VBIDE.VBComponent
    For Each c In proj.VBComponents
        If InStr(1, beforeSet, "|" & c.Name & "|", vbTextCompare) = 0 Then
            NewComponentName = c.Name
            Exit Function
        End If
    Next
End Function

Private Function NextFreeComponentName(ByVal proj As VBIDE.VBProject, ByVal stem As String) As String
    Dim n As Long
    Dim c As VBIDE.VBComponent
    Dim candidate As String
    Dim taken As Boolean
    n = 1
    Do
        candidate = stem & CStr(n)
        taken = False
        For Each c In proj.VBComponents
            If StrComp(c.Name, candidate, vbTextCompare) = 0 Then
                taken = True
                Exit For
            End If
        Next
        If Not taken Then
            NextFreeComponentName = candidate
            Exit Function
        End If
        n = n + 1
    Loop
End Function

' LAST-RESORT synthesized .frm text (TASK-P1-084 kept it as a fallback: the IDE
' route above is the normal path now). Two defects fixed here:
'   (1) LinkTopic was hard-coded "Form1" regardless of compName - VB6 writes the
'       form's OWN name (measured: the IDE's new Form1.frm has LinkTopic="Form1";
'       our own SettingsForm.frm still carries the bogus "Form1" from this line);
'   (2) the geometry was a hard-coded approximation the IDE never produces
'       (3090/60/450/4680) - it is now the IDE's REAL default for a new form,
'       measured byte-for-byte from an IDE-generated Form1.frm:
'       ClientHeight=3015 ClientLeft=120 ClientTop=465 ClientWidth=4560.
Private Function FormFileText(ByVal compName As String) As String
    Dim s As String
    s = "VERSION 5.00" & vbCrLf
    s = s & "Begin VB.Form " & compName & " " & vbCrLf
    s = s & "   Caption         =   """ & compName & """" & vbCrLf
    s = s & "   ClientHeight    =   3015" & vbCrLf
    s = s & "   ClientLeft      =   120" & vbCrLf
    s = s & "   ClientTop       =   465" & vbCrLf
    s = s & "   ClientWidth     =   4560" & vbCrLf
    s = s & "   LinkTopic       =   """ & compName & """" & vbCrLf
    s = s & "   ScaleHeight     =   3015" & vbCrLf
    s = s & "   ScaleWidth      =   4560" & vbCrLf
    s = s & "   StartUpPosition =   3  'Windows Default" & vbCrLf
    s = s & "End" & vbCrLf
    s = s & "Attribute VB_Name = """ & compName & """" & vbCrLf
    s = s & "Attribute VB_GlobalNameSpace = False" & vbCrLf
    s = s & "Attribute VB_Creatable = False" & vbCrLf
    s = s & "Attribute VB_PredeclaredId = True" & vbCrLf
    s = s & "Attribute VB_Exposed = False" & vbCrLf
    FormFileText = s
End Function

Private Function CtlJoinPath(ByVal dir As String, ByVal leaf As String) As String
    If Len(dir) = 0 Then
        CtlJoinPath = leaf
    ElseIf Right$(dir, 1) = "\" Then
        CtlJoinPath = dir & leaf
    Else
        CtlJoinPath = dir & "\" & leaf
    End If
End Function

Private Function CtlFileExists(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CtlFileExists = fso.FileExists(path)
End Function

Private Sub CtlWriteFile(ByVal path As String, ByVal text As String)
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Output As #fnum
    Print #fnum, text;
    Close #fnum
End Sub

Private Function ToolRemoveComponent(ByVal name As String) As String
    ' TASK-P1-121 item 5 (G-2), VARIANT (a) = REPORT, DO NOT DELETE.
    ' P4-104 measured that VBComponents.Remove writes <ProjectDir>\<name><ext> on its way
    ' out while the receipt said only "removed", so a caller could not tell that a file had
    ' appeared. Variant (b) (delete that file) would make the frozen description's
    ' "the IDE will ask about the file" false and would need a second description change,
    ' so this card takes variant (a): the receipt NAMES what was left behind and the
    ' additive sideEffects block repeats it. Nothing is deleted.
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim boundFile As String
    boundFile = ""
    On Error Resume Next
    boundFile = comp.FileNames(1)
    If Err.Number <> 0 Then
        Err.Clear
        boundFile = ""
    End If
    On Error GoTo 0
    Dim proj As VBIDE.VBProject
    Set proj = GetProject()
    proj.VBComponents.Remove comp
    Dim afterFile As String
    afterFile = ""
    On Error Resume Next
    If Len(boundFile) > 0 Then
        If Dir$(boundFile) <> "" Then afterFile = boundFile
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("removed") = name
    root("removedFile") = afterFile
    If Len(afterFile) > 0 Then
        root("removedFileNote") = "VBComponents.Remove left this file in the project directory; it was NOT deleted (the component's source is preserved)."
    End If
    Dim writes As Object
    Set writes = New Collection
    ManifestAdd writes, afterFile
    ManifestMergeUndo writes
    Set root("projectDirWrites") = writes
    SeNoteRemoved name
    ToolRemoveComponent = JsonStringify(root)
End Function

Private Function ToolImportComponent(ByVal path As String, ByVal mode As String) As String
    ' TASK-P1-056: importing must NEVER write back to the file the caller gave us.
    ' VBComponents.AddFile binds the new component's file to that path, so a later save
    ' of the component overwrites the caller's source - this is how a real source file
    ' was silently replaced by a 31-byte "Attribute VB_Name = Module1" stub. We now
    ' re-point the component at a PROJECT-LOCAL copy immediately, verify the component
    ' carries the file's real code, snapshot the source only when it is WRITABLE, and
    ' verify the caller's file is byte-identical afterwards (restoring it if not).
    ' A READ-ONLY source stays importable (ruling A) and leaves no .bak behind.
    ' TASK-P1-060 tightens this to a testable INVARIANT: on return NO component may still be
    ' bound to the caller's path (see (b) below), because a merely-later save would otherwise
    ' destroy the source after this method has already verified it.
    Dim full As String
    full = ResolvePath(path)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(full) Then Err.Raise vbObjectError + 48, "CodeTools", "File not found: " & full
    ' ============ TASK-P1-071 (DATA-LOSS level): refuse BEFORE ANY side effect ==========
    ' Measured on AFAC95AF (TASK-P4-043, reproduced TWICE): importing a .bas whose
    ' "Attribute VB_Name" equals an EXISTING component does not raise - VB6 creates an extra
    ' default-named component (Module1) still bound to the caller's path, SaveAs then writes
    ' <ProjectDir>\<name><ext> and REBINDS the pre-existing component to that file, the
    ' component set collapses (213 -> 8..9) and the IDE exits by itself ~63 s later.
    ' There was NO same-name check anywhere in this source. The 60059 that P1-063 saw is VB6's
    ' OWN AddFile error for a path that is ALREADY a project member: it fires only for a repeat
    ' import of the SAME path and never covered the same-NAME / different-path case.
    ' Both refusals sit HERE - after ResolvePath/FileExists and BEFORE the writable-source
    ' snapshot (FreeBackupPath below), before AddFile and before SaveAs - so a refusal leaves
    ' NOTHING behind: no <file>.mcpimport.bak next to the caller's file, no new file in the
    ' project directory, no component added and no component rebound.
    Dim pj0 As VBIDE.VBProject
    Set pj0 = GetProject()
    ' TASK-P1-077: a project-LOCAL copy needs a project DIRECTORY. When the IDE holds a project
    ' that has never been saved, ProjectDir() is empty; FreeImportPath used to fall back to
    ' App.Path and the copy landed in the add-in's own directory (product directory). Refuse
    ' HERE - after ResolvePath/FileExists and BEFORE the writable-source snapshot, AddFile and
    ' SaveAs - so the refusal has no side effect at all.
    Dim pd0 As String
    pd0 = ProjectDir()
    If Len(pd0) = 0 Then
        Err.Raise vbObjectError + 98, "CodeTools", "MCPREASON:project_dir_unavailable|" & _
            "save the project first (save_project with a path) so the import has a project directory for its project-local copy|" & _
            "the open project has no file on disk, so there is no project directory to put the project-local copy in. Refused BEFORE the source snapshot, before AddFile and before SaveAs: nothing is written next to the source file, into the project directory or into the add-in's own directory."
    End If
    Dim declaredName As String
    declaredName = DeclaredComponentName(full)
    If Len(declaredName) = 0 Then
        ' No Attribute VB_Name in the head (e.g. a code-only .bas): VB6 then names the new
        ' component after the FILE, so the file's base name is the name that can collide.
        Dim bp0 As Long
        bp0 = InStrRev(full, "\")
        declaredName = Mid$(full, bp0 + 1)
        bp0 = InStrRev(declaredName, ".")
        If bp0 > 1 Then declaredName = Left$(declaredName, bp0 - 1)
    End If
    Dim ec As VBIDE.VBComponent
    Dim ecName As String, pathOwner As String, ef As String
    ecName = ""
    pathOwner = ""
    For Each ec In pj0.VBComponents
        MaybePump
        If Len(ecName) = 0 And Len(declaredName) > 0 Then
            If LCase$(ec.Name) = LCase$(declaredName) Then ecName = ec.Name
        End If
        If Len(pathOwner) = 0 Then
            ef = ComponentFilePath(ec)
            If Len(ef) > 0 Then
                If LCase$(ef) = LCase$(full) Then pathOwner = ec.Name
            End If
        End If
        If Len(ecName) > 0 And Len(pathOwner) > 0 Then Exit For
    Next
    If Len(pathOwner) > 0 Then
        Err.Raise vbObjectError + 97, "CodeTools", "MCPREASON:import_path_is_member|" & pathOwner & "|" & "this file is already the source of component '" & pathOwner & "'; edit that component with replace_code / edit_by_anchor / apply_edits instead of importing the file again"
    End If
    If Len(ecName) > 0 Then
        Err.Raise vbObjectError + 96, "CodeTools", "MCPREASON:name_collision|" & ecName & "|" & "a component named '" & ecName & "' already exists and the incoming file declares Attribute VB_Name = """ & declaredName & """. Rename one of them, delete the existing component first, or edit the existing component with replace_code / edit_by_anchor / apply_edits"
    End If
    ' ============ end TASK-P1-071 pre-check ============================================
    ' TASK-P1-056 ruling A: a READ-ONLY source is LEGAL and must stay importable (CD,
    ' archive, read-only share). We never write it anyway. A snapshot is made ONLY when the
    ' source is WRITABLE - only then could a restore ever be needed - so a read-only medium
    ' is never asked to hold a .bak and no extra file is left behind there.
    Dim shaBefore As String
    shaBefore = Build.FileSha256(full)
    Dim canWrite As Boolean, backup As String, haveBackup As Boolean
    canWrite = ((fso.GetFile(full).Attributes And 1) = 0)
    haveBackup = False
    backup = ""
    If canWrite Then
        backup = FreeBackupPath(fso, full)
        On Error Resume Next
        fso.CopyFile full, backup, True
        If Err.Number = 0 Then haveBackup = True Else Err.Clear
        On Error GoTo 0
    End If
    Dim comp As VBIDE.VBComponent
    Dim proj As VBIDE.VBProject
    Set proj = GetProject()
    ' TASK-P1-071: AddFile itself can refuse (VB6 error 60059: the path is ALREADY a project
    ' member - the P1-063 observation). That refusal happens AFTER the writable-source snapshot,
    ' so the snapshot file used to be left behind next to the caller's file. Drop it here and
    ' re-raise the ORIGINAL VB error, so the envelope stays exactly the unexpected_error shape
    ' it always was while the refusal stops leaving a residue.
    Dim addErrNum As Long, addErrDesc As String
    addErrNum = 0
    addErrDesc = ""
    On Error Resume Next
    If LCase$(mode) = "template" Then
        Set comp = proj.VBComponents.AddFromTemplate(full)
    Else
        Set comp = proj.VBComponents.AddFile(full, False)
    End If
    addErrNum = Err.Number
    addErrDesc = Err.Description
    Err.Clear
    On Error GoTo 0
    If addErrNum <> 0 Then
        If Len(backup) > 0 Then
            On Error Resume Next
            If fso.FileExists(backup) Then fso.DeleteFile backup, True
            On Error GoTo 0
        End If
        Err.Raise addErrNum, "CodeTools", addErrDesc
    End If
    ' (b) TASK-P1-060 INVARIANT: on return NO component may still be bound to the caller's
    ' source path. TASK-P1-056 rebound only the component AddFile returned; when the project
    ' already has a component with the SAME NAME (from a different file) VB6 creates an
    ' ADDITIONAL default-named component (Module1) that keeps the binding and later overwrites
    ' the source - AFTER this method verified the sha. Rebind EVERY bound component, then VERIFY.
    ' The whole block runs under On Error Resume Next: a runtime error from a VBIDE call must
    ' never be able to skip the verification (that is exactly how error 5 escaped before).
    Dim newPath As String, rebindErr As String, loopErr As String
    rebindErr = ""
    loopErr = ""
    newPath = ""
    Dim rebound As Long, stillBound As Long, bNames As String
    rebound = 0
    stillBound = 0
    bNames = ""
    Dim createdCnt As Long, reboundList As String, saveAsErr As String
    createdCnt = 0
    reboundList = ""
    Dim pass As Long, c2 As VBIDE.VBComponent, fn As String, rp As String, j As Long
    ' TASK-P1-071: every file THIS call creates via SaveAs is remembered here so the refusal
    ' path can remove it again (the reject path used to restore only the CALLER's file and left
    ' <ProjectDir>\<name><ext> on disk - exactly the stray file TASK-P4-043 found).
    Dim madePaths As Collection
    Set madePaths = New Collection
    On Error Resume Next
    For pass = 1 To 2
        For j = 1 To proj.VBComponents.Count
            MaybePump
            Err.Clear
            Set c2 = proj.VBComponents(j)
            fn = ""
            If Not c2 Is Nothing Then fn = c2.FileNames(1)
            If Err.Number <> 0 Then
                If Len(loopErr) = 0 Then loopErr = "scan " & CStr(Err.Number) & ": " & Err.Description
                Err.Clear
            End If
            If Len(fn) > 0 Then
                If LCase$(fn) = LCase$(full) Then
                    If pass = 1 Then createdCnt = createdCnt + 1
                    rp = FreeImportPath(fso, c2.Name, full)
                    If Err.Number <> 0 Then
                        If Len(loopErr) = 0 Then loopErr = "path " & CStr(Err.Number) & ": " & Err.Description
                        Err.Clear
                        rp = ""
                    End If
                    If Len(rp) > 0 Then
                        c2.SaveAs rp
                        saveAsErr = ""
                        If Err.Number <> 0 Then
                            saveAsErr = CStr(Err.Number) & ": " & Err.Description
                            Err.Clear
                        End If
                        ' TASK-P1-060: SaveAs can still raise VB error 5 AFTER it has in fact
                        ' re-pathed the component (measured on every fixture: code and 1.2 MB
                        ' alike), so the SUCCESS signal is the BINDING itself, not Err. A
                        ' component that truly could not be re-pathed is still bound to the
                        ' caller path, and the invariant scan below refuses on stillBound > 0.
                        fn = ""
                        Err.Clear
                        fn = c2.FileNames(1)
                        If Err.Number <> 0 Then Err.Clear
                        If Len(fn) > 0 And LCase$(fn) <> LCase$(full) Then
                            rebound = rebound + 1
                            madePaths.Add fn
                            If Len(reboundList) < 400 Then reboundList = reboundList & c2.Name & "=" & fn & "; "
                            If Not comp Is Nothing Then
                                If LCase$(c2.Name) = LCase$(comp.Name) Then newPath = rp
                            End If
                        ElseIf Len(rebindErr) = 0 Then
                            rebindErr = c2.Name & ": the binding still points at the source path"
                            If Len(saveAsErr) > 0 Then rebindErr = rebindErr & " (SaveAs " & saveAsErr & ")"
                        End If
                    End If
                End If
            End If
        Next
    Next
    ' the invariant scan must ALWAYS run
    For j = 1 To proj.VBComponents.Count
        MaybePump
        Err.Clear
        Set c2 = proj.VBComponents(j)
        fn = ""
        If Not c2 Is Nothing Then fn = c2.FileNames(1)
        If Err.Number <> 0 Then Err.Clear
        If Len(fn) > 0 Then
            If LCase$(fn) = LCase$(full) Then
                stillBound = stillBound + 1
                If Len(bNames) < 120 Then bNames = bNames & c2.Name & " "
            End If
        End If
    Next
    On Error GoTo 0
    If stillBound > 0 Then
        If haveBackup Then
            On Error Resume Next
            fso.CopyFile backup, full, True
            On Error GoTo 0
        End If
        If Len(backup) > 0 Then
            On Error Resume Next
            If fso.FileExists(backup) Then fso.DeleteFile backup, True
            On Error GoTo 0
        End If
        ' TASK-P1-071: this refusal is reached AFTER SaveAs, so the file(s) this call created in
        ' the PROJECT DIRECTORY must go too, otherwise a refusal still leaves a stray
        ' <ProjectDir>\<name><ext> behind (P4-046 static finding; P4-043 saw it on disk).
        Dim mk As Long
        On Error Resume Next
        For mk = 1 To madePaths.Count
            If fso.FileExists(madePaths(mk)) Then fso.DeleteFile madePaths(mk), True
        Next
        On Error GoTo 0
        Err.Raise vbObjectError + 54, "CodeTools", "import refused: " & CStr(stillBound) & " component(s) are still bound to the source path (" & bNames & "); rebind failed (" & rebindErr & ") " & loopErr & ". The source was restored from the snapshot and nothing was left bound."
    End If
    If Len(newPath) = 0 Then
        On Error Resume Next
        If Not comp Is Nothing Then newPath = comp.FileNames(1)
        On Error GoTo 0
    End If
    ' (a) the component must carry the file's REAL code (AddFile can mis-parse a
    ' non-ASCII file and leave an empty module). TASK-P1-060: (a) must NEVER be able to
    ' abort this function. CompText does "cm.Lines(1, cm.CountOfLines)", and the
    ' default-named module VB6 creates on a NAME collision is EMPTY (CountOfLines = 0),
    ' so "Lines(1, 0)" raises VB error 5. That raise used to escape from here BEFORE (c)
    ' ran, so the sha verify + restore never executed and the clobbered source stayed
    ' clobbered. A read failure is now DEFERRED to the end (after (c) has restored).
    Dim payload As String, wantCode As String, gotCode As String
    Dim codeFromFile As Boolean, readFailed As Boolean
    codeFromFile = True
    readFailed = False
    On Error Resume Next
    payload = Build.ReadLogAnsi(full)
    If Err.Number <> 0 Then
        readFailed = True
        Err.Clear
    End If
    On Error GoTo 0
    If Not readFailed Then
        On Error Resume Next
        ' TASK-P1-121 item 1 (S1): the text handed to SetCompTextMinimal is now the
        ' STRUCTURAL head-stripped body, not NormalizeCode(payload). NormalizeCode only
        ' removes "Attribute " lines, so the .cls VERSION/BEGIN..END designer block and
        ' the .frm Begin..End block used to be written INTO the code pane (P4-097/099
        ' measured it on both .cls and code-bearing .frm). The COMPARISON still uses
        ' NormalizeCode, so the caliber of the comparison itself is unchanged.
        Dim wantCaliber As String
        wantCaliber = "unknown"
        wantCode = StrippedDiskText(payload, SafeCompText(comp.Name), True, wantCaliber)
        gotCode = NormalizeCode(CompText(comp.Name))
        If Err.Number <> 0 Then
            ' could not compare (empty module, ...): report "not proven to come from the
            ' file" instead of aborting the function before (c) ran.
            Err.Clear
            codeFromFile = False
        ElseIf wantCaliber <> "ok" Then
            ' TASK-P1-121 item 3 failure safety, declared by the P4-106 spec (section 1.3
            ' and section 5): the split could not be reconciled with the pane (empty pane,
            ' head-only file). We do not guess - nothing is written, the receipt says the
            ' component is NOT proven to carry the file's code, and (c) still verifies the
            ' caller's file. Declared behaviour change: an empty pane is no longer
            ' "repaired" with the file's designer head.
            codeFromFile = False
        Else
            If NormalizeCode(wantCode) <> gotCode Then
                ' TASK-P1-059: incremental write instead of delete-all + insert-all.
                SetCompTextMinimal comp.Name, CompText(comp.Name), wantCode
                If Err.Number <> 0 Then Err.Clear
                codeFromFile = False
            End If
        End If
        On Error GoTo 0
    End If
    ' (c) the caller's file must be byte-identical; restore from the snapshot if not
    Dim shaAfter As String, restored As Boolean
    restored = False
    shaAfter = Build.FileSha256(full)
    If shaAfter <> shaBefore Then
        If haveBackup Then
            On Error Resume Next
            fso.CopyFile backup, full, True
            On Error GoTo 0
            shaAfter = Build.FileSha256(full)
            restored = True
        End If
    End If
    If haveBackup Then
        On Error Resume Next
        If fso.FileExists(backup) Then fso.DeleteFile backup, True
        On Error GoTo 0
    End If
    ' TASK-P1-060: this is the ONLY hard error, and it is raised HERE - after (c) has
    ' verified/restored the caller's file and dropped the snapshot - so even a source
    ' that cannot be read leaves no clobbered file and no .bak behind.
    If readFailed Then
        Err.Raise vbObjectError + 52, "CodeTools", "cannot READ the source file (it may be missing, locked, or not text): " & full
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("imported") = comp.Name
    SeNoteCreated comp.Name
    root("type") = ComponentTypeName(comp.Type)
    root("sourceFile") = full
    root("sourceFileSha256Before") = shaBefore
    root("sourceFileSha256After") = shaAfter
    root("sourceFileUnchanged") = (shaAfter = shaBefore)
    root("sourceFileRestored") = restored
    root("sourceWritable") = canWrite
    If Len(backup) > 0 Then root("snapshotFile") = backup
    root("componentFile") = newPath
    root("componentsRebound") = rebound
    root("componentsStillBound") = stillBound
    root("componentsCreated") = createdCnt
    ' TASK-P1-077: the write manifest is ALWAYS present (empty array when the call wrote no
    ' file) so that a caller can decide from the receipt alone instead of guessing what the
    ' tool left behind. Every path here is a SaveAs target that rebinds the new component off
    ' the caller's file, so by construction it lives INSIDE the project directory.
    Dim pdw As Object
    Set pdw = New Collection
    Dim mk2 As Long
    On Error Resume Next
    For mk2 = 1 To madePaths.Count
        ManifestAdd pdw, madePaths(mk2)
    Next
    On Error GoTo 0
    ManifestMergeUndo pdw
    Set root("projectDirWrites") = pdw
    If Len(reboundList) > 0 Then root("componentsReboundFiles") = reboundList
    If Len(rebindErr) > 0 Then root("componentRebindError") = rebindErr
    root("codeFromFile") = codeFromFile
    root("note") = "import is READ-ONLY with respect to the file you passed: the component is bound to a project-local copy and the source file's sha256 is verified (and restored from a snapshot) afterwards"
    ToolImportComponent = JsonStringify(root)
End Function

' TASK-P1-056: a free project-local path for a freshly imported component, keeping the
' caller's extension. Never returns the caller's own path.
' TASK-P1-056 ruling B: never overwrite a user's existing <file>.mcpimport.bak; take the
' next numbered name instead and say so in the receipt.
Private Function FreeBackupPath(ByVal fso As Object, ByVal originalFull As String) As String
    Dim base As String, cand As String, n As Long
    base = originalFull & ".mcpimport.bak"
    cand = base
    n = 0
    Do While fso.FileExists(cand)
        n = n + 1
        cand = base & "." & CStr(n)
    Loop
    FreeBackupPath = cand
End Function

' TASK-P1-057: write ONLY the changed range instead of SetCompText's delete-the-whole-
' module + insert-the-whole-module. SetCompText was measured at ~15.7 s on a 27,000-line /
' 1.2 MB module (two whole-module reformats), while replace_code's one-range write is 0 ms.
' The common prefix/suffix are found with the same algorithm ChangedRegion uses, so the
' resulting pane text is identical to writing the whole text - just two CM calls at most.
Private Function SetCompTextMinimal(ByVal name As String, ByVal before As String, ByVal after As String, Optional ByVal snapshotAlreadyTaken As Boolean = False) As Long
    ' TASK-P1-068: a small edit far from the other end of the module used to collapse
    ' into ONE block spanning 'first difference .. last difference', and a single CM
    ' block costs t ~ k*s^2 (measured 8,438 ms for 26,999 lines). The change is now
    ' decomposed into line-level hunks and each hunk is written on its own (<=1
    ' DeleteLines + <=1 InsertLines). The resulting pane is byte-identical.
    ' Invariants: hunk count <= MAX_HUNKS; no hunk larger than the single-block span;
    ' anything else falls back to the original single minimal block and the receipt
    ' says so (hunks=1).
    ' TASK-P1-078: the per-call receipt state is reset FIRST, so the two early exits
    ' below (missing component / missing code module) can no longer leave the PREVIOUS
    ' call's hunks and write phases readable by the caller.
    m_lastWriteHunks = 0
    m_lastDiffMs = 0
    m_lastJoinMs = 0
    m_lastDelMs = 0
    m_lastInsMs = 0
    m_lastMaxBlock = 0
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Exit Function
    Dim writes As Long
    writes = 0
    Dim tPh As Long
    tPh = GetTickCount()
    Dim a() As String, b() As String, na As Long, nb As Long, pre As Long, suf As Long, i As Long
    a = TextLines(before)
    b = TextLines(after)
    na = UBound(a) + 1
    nb = UBound(b) + 1
    pre = 0
    Do While pre < na And pre < nb
        MaybePump
        If a(pre) <> b(pre) Then Exit Do
        pre = pre + 1
    Loop
    suf = 0
    Do While suf < (na - pre) And suf < (nb - pre)
        MaybePump
        If a(na - 1 - suf) <> b(nb - 1 - suf) Then Exit Do
        suf = suf + 1
    Loop
    Dim delCount As Long, insCount As Long
    delCount = na - pre - suf
    If delCount < 0 Then delCount = 0
    insCount = nb - pre - suf
    If insCount < 0 Then insCount = 0
    m_lastJoinMs = m_lastJoinMs + (GetTickCount() - tPh)
    If delCount = 0 And insCount = 0 Then Exit Function
    ' TASK-P1-076 (P8): snapshot the PRE-WRITE pane text just before the change lands, so
    ' undo_edit has something to restore. Never fatal, and disabled together with every
    ' other project-directory write (ProjectDirWritesKilled).
    ' TASK-P1-087: an opt-in caller may have taken the snapshot itself (it is the only way
    ' the raw replace_code branches get one); in that case this funnel must NOT write a
    ' second, identical file. Default False keeps every existing caller byte-for-byte.
    If Not snapshotAlreadyTaken Then UndoSnapshot name, before
    Dim hs() As Long, nh As Long, h As Long
    Dim ok As Boolean
    nh = 0
    ReDim hs(0 To MAX_HUNKS * 4 - 1)
    tPh = GetTickCount()
    ok = HunkDiff(a, b, pre, na - suf, pre, nb - suf, hs, nh)
    If ok And nh >= 2 Then
        ' validate EVERY hunk BEFORE the first write: once a hunk is applied a
        ' fallback could no longer reproduce the original pane.
        Dim ax As Long, ac As Long, bx As Long, bc As Long
        For h = 0 To nh - 1
            ax = hs(h * 4 + 0)
            ac = hs(h * 4 + 1)
            bx = hs(h * 4 + 2)
            bc = hs(h * 4 + 3)
            If ac > delCount Then ok = False
            If bc > insCount Then ok = False
            If ax < 0 Or ax + ac > na Then ok = False
            If bx < 0 Or bx + bc > nb Then ok = False
        Next
    Else
        ok = False
    End If
    m_lastDiffMs = m_lastDiffMs + (GetTickCount() - tPh)
    If ok Then
        For h = nh - 1 To 0 Step -1
            ax = hs(h * 4 + 0)
            ac = hs(h * 4 + 1)
            bx = hs(h * 4 + 2)
            bc = hs(h * 4 + 3)
            If ac > 0 Then
                tPh = GetTickCount()
                cm.DeleteLines ax + 1, ac
                m_lastDelMs = m_lastDelMs + (GetTickCount() - tPh)
                writes = writes + 1
                If ac > m_lastMaxBlock Then m_lastMaxBlock = ac
            End If
            If bc > 0 Then
                Dim blk() As String
                ReDim blk(0 To bc - 1)
                Dim q As Long
                tPh = GetTickCount()
                For q = 0 To bc - 1
                    MaybePump
                    blk(q) = b(bx + q)
                Next
                Dim blkText As String
                blkText = Join(blk, vbCrLf)
                m_lastJoinMs = m_lastJoinMs + (GetTickCount() - tPh)
                tPh = GetTickCount()
                cm.InsertLines ax + 1, blkText
                m_lastInsMs = m_lastInsMs + (GetTickCount() - tPh)
                writes = writes + 1
                If bc > m_lastMaxBlock Then m_lastMaxBlock = bc
            End If
            MaybePump
        Next
        m_lastWriteHunks = nh
        SetCompTextMinimal = writes
        Exit Function
    End If
    ' ---- fallback: the pre-TASK-P1-068 single minimal block (writeCalls <= 2) ----
    writes = 0
    m_lastWriteHunks = 1
    If delCount > 0 Then
        tPh = GetTickCount()
        cm.DeleteLines pre + 1, delCount
        m_lastDelMs = m_lastDelMs + (GetTickCount() - tPh)
        writes = writes + 1
        If delCount > m_lastMaxBlock Then m_lastMaxBlock = delCount
    End If
    If insCount > 0 Then
        Dim blkF() As String
        ReDim blkF(0 To insCount - 1)
        tPh = GetTickCount()
        For i = 0 To insCount - 1
            MaybePump
            blkF(i) = b(pre + i)
        Next
        Dim blkFText As String
        blkFText = Join(blkF, vbCrLf)
        m_lastJoinMs = m_lastJoinMs + (GetTickCount() - tPh)
        tPh = GetTickCount()
        cm.InsertLines pre + 1, blkFText
        m_lastInsMs = m_lastInsMs + (GetTickCount() - tPh)
        writes = writes + 1
        If insCount > m_lastMaxBlock Then m_lastMaxBlock = insCount
    End If
    SetCompTextMinimal = writes
End Function

' TASK-P1-070: publish the per-phase timings of the last write into a receipt.
' Pure addition: writeMs/writeCalls/hunks keep their exact meaning. phaseNote
' spells out what each interval covers so no cause is read into a measurement.
Private Sub AddWritePhases(ByRef d As Object)
    d("diffMs") = m_lastDiffMs
    d("joinMs") = m_lastJoinMs
    d("delMs") = m_lastDelMs
    d("insMs") = m_lastInsMs
    d("maxBlockLines") = m_lastMaxBlock
    d("phaseNote") = "measured intervals of the LAST write of THIS call: diffMs=hunk generation+validation; joinMs=the rest of our own preparation (line split, common prefix/suffix scan, block build, Join); delMs/insMs=time inside the CodeModule calls; maxBlockLines=largest number of lines in one CodeModule call. writeMs/writeCalls/hunks keep their existing meaning"
End Sub

' TASK-P1-068: recursive line diff. Regions are trimmed by their common prefix and
' suffix first; a region that still differs is split at a line that is UNIQUE on
' both sides (the longest such line wins), then the two halves are diffed again.
' A region with no anchor becomes one hunk. Returns False at the hunk cap so the
' caller falls back to the single-block writer.
Private Function HunkDiff(ByRef a() As String, ByRef b() As String, ByVal alo As Long, ByVal ahi As Long, _
                          ByVal blo As Long, ByVal bhi As Long, ByRef hs() As Long, ByRef nh As Long) As Boolean
    Do While alo < ahi And blo < bhi
        MaybePump
        If a(alo) <> b(blo) Then Exit Do
        alo = alo + 1
        blo = blo + 1
    Loop
    Do While ahi > alo And bhi > blo
        MaybePump
        If a(ahi - 1) <> b(bhi - 1) Then Exit Do
        ahi = ahi - 1
        bhi = bhi - 1
    Loop
    If alo = ahi And blo = bhi Then
        HunkDiff = True
        Exit Function
    End If
    Dim ai As Long, bi As Long
    If FindAnchor(a, b, alo, ahi, blo, bhi, ai, bi) Then
        If Not HunkDiff(a, b, alo, ai, blo, bi, hs, nh) Then Exit Function
        If Not HunkDiff(a, b, ai + 1, ahi, bi + 1, bhi, hs, nh) Then Exit Function
        HunkDiff = True
        Exit Function
    End If
    If nh >= MAX_HUNKS Then Exit Function
    hs(nh * 4 + 0) = alo
    hs(nh * 4 + 1) = ahi - alo
    hs(nh * 4 + 2) = blo
    hs(nh * 4 + 3) = bhi - blo
    nh = nh + 1
    HunkDiff = True
End Function

' TASK-P1-068: the longest line that occurs exactly once on BOTH sides of the region.
Private Function FindAnchor(ByRef a() As String, ByRef b() As String, ByVal alo As Long, ByVal ahi As Long, _
                            ByVal blo As Long, ByVal bhi As Long, ByRef ai As Long, ByRef bi As Long) As Boolean
    Dim ca As Object, cb As Object, pos As Object
    Set ca = CreateObject("Scripting.Dictionary")
    Set cb = CreateObject("Scripting.Dictionary")
    Set pos = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    For i = alo To ahi - 1
        MaybePump
        k = a(i)
        If Len(k) > 0 Then
            If ca.Exists(k) Then ca(k) = CLng(ca(k)) + 1 Else ca.Add k, 1
        End If
    Next
    For i = blo To bhi - 1
        MaybePump
        k = b(i)
        If Len(k) > 0 Then
            If cb.Exists(k) Then cb(k) = CLng(cb(k)) + 1 Else cb.Add k, 1
            pos(k) = i
        End If
    Next
    Dim best As Long, bestA As Long
    best = 0
    bestA = -1
    bi = -1
    For i = alo To ahi - 1
        MaybePump
        k = a(i)
        If Len(k) > best Then
            If ca.Exists(k) Then
                If CLng(ca(k)) = 1 Then
                    If cb.Exists(k) Then
                        If CLng(cb(k)) = 1 Then
                            best = Len(k)
                            bestA = i
                            bi = CLng(pos(k))
                        End If
                    End If
                End If
            End If
        End If
    Next
    If bestA >= 0 Then
        ai = bestA
        FindAnchor = True
    End If
End Function

' TASK-P1-071: the component name a file DECLARES, read from the file head, "" when absent.
' .bas/.cls keep "Attribute VB_Name" in the HEADER; a designer file (.frm/.ctl/.dsr) carries
' its name on the "Begin VB.Form <Name>" line instead - its Attribute block lives at the TAIL
' of the file, so a header-only scan would silently miss every form.
Private Function DeclaredComponentName(ByVal full As String) As String
    Dim f As Integer, i As Long, ln As String, s As String
    Dim q1 As Long, q2 As Long
    DeclaredComponentName = ""
    On Error Resume Next
    f = FreeFile
    Open full For Input As #f
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    For i = 1 To 60
        If EOF(f) Then Exit For
        Line Input #f, ln
        If InStr(1, ln, "Attribute VB_Name", vbTextCompare) > 0 Then
            q1 = InStr(ln, Chr$(34))
            If q1 > 0 Then
                q2 = InStr(q1 + 1, ln, Chr$(34))
                If q2 > q1 Then DeclaredComponentName = Mid$(ln, q1 + 1, q2 - q1 - 1)
            End If
            Exit For
        End If
        If InStr(1, ln, "Begin VB.", vbTextCompare) > 0 Then
            s = Trim$(ln)
            q1 = InStrRev(s, " ")
            If q1 > 0 Then s = Trim$(Mid$(s, q1 + 1)) Else s = ""
            If Len(s) > 0 And InStr(s, "=") = 0 And InStr(s, Chr$(34)) = 0 Then DeclaredComponentName = s
            Exit For
        End If
    Next
    Close #f
    On Error GoTo 0
End Function

' TASK-P1-071: the file a component is bound to, or "" when it cannot be read. FileNames(1)
' raises for some component kinds and in a busy IDE; a failed read must never abort the scan.
Private Function ComponentFilePath(ByVal comp As VBIDE.VBComponent) As String
    Dim s As String
    ComponentFilePath = ""
    If comp Is Nothing Then Exit Function
    On Error Resume Next
    s = comp.FileNames(1)
    If Err.Number <> 0 Then
        Err.Clear
        s = ""
    End If
    On Error GoTo 0
    ComponentFilePath = s
End Function

Private Function FreeImportPath(ByVal fso As Object, ByVal compName As String, ByVal originalFull As String) As String
    ' TASK-P1-077: there is NO App.Path fallback here any more. ProjectDir() is empty when the
    ' IDE holds no project OR a project that has never been saved, and the old
    ' "If Len(d) = 0 Then d = App.Path" then wrote the import's project-local copy into the
    ' ADD-IN'S OWN directory - measured by TASK-P1-071 as "<product dir>\Module1.bas" (183 B,
    ' archived at 内部工作区). The caller refuses before any
    ' side effect (see the project-dir guard in ToolImportComponent); returning "" here keeps
    ' the SaveAs step unreachable even if that guard were ever bypassed, because every caller
    ' tests Len(rp) > 0.
    Dim d As String, ext As String, p As Long, base As String, cand As String, n As Long
    d = ProjectDir()
    If Len(d) = 0 Then Exit Function
    p = InStrRev(originalFull, ".")
    If p > 0 Then ext = Mid$(originalFull, p) Else ext = ".bas"
    base = d & "\" & compName
    cand = base & ext
    n = 1
    Do While fso.FileExists(cand)
        n = n + 1
        cand = base & "_" & CStr(n) & ext
    Loop
    FreeImportPath = cand
End Function

' TASK-P1-077: operator switch for writes that put a file into the project directory on behalf
' of the caller. It is an environment switch and NOT a tool parameter, so the tool face and
' every inputSchema stay untouched (no re-freeze); it is read in the VB6 process because the
' add-in runs in-process in VB6.EXE.
' Semantics: unset / 1 / true / yes / on = allowed (the caller's explicit request is the
' opt-in: only "add_component {type:\"form\"}" can ever reach the writing path); 0 / false /
' no / off = REFUSED, which is what a read-only or audit setup wants.
Private Function ProjectDirWritesKilled() As Boolean
    Dim v As String
    v = LCase$(Trim$(Environ$("VB6MCP_ALLOW_PROJECTDIR_WRITES")))
    ProjectDirWritesKilled = (v = "0" Or v = "false" Or v = "no" Or v = "off")
End Function

' TASK-P1-077: build the "projectDirWrites" write manifest. Absolute paths, write order,
' de-duplicated (SaveAs is retried in a second pass and a retry must not appear twice).
Private Sub ManifestAdd(ByVal coll As Object, ByVal path As String)
    Dim i As Long
    If Len(path) = 0 Then Exit Sub
    If coll Is Nothing Then Exit Sub
    For i = 1 To coll.Count
        If LCase$(coll(i)) = LCase$(path) Then Exit Sub
    Next
    coll.Add path
End Sub

' TASK-P1-085: merge the _undo snapshot files written during THIS call into a receipt's
' projectDirWrites manifest. Order preserved; ManifestAdd de-duplicates.
Private Sub ManifestMergeUndo(ByVal coll As Object)
    Dim i As Long
    On Error Resume Next
    If coll Is Nothing Then Exit Sub
    If m_dirWrites Is Nothing Then Exit Sub
    For i = 1 To m_dirWrites.Count
        ManifestAdd coll, CStr(m_dirWrites(i))
    Next
    Err.Clear
End Sub

' ===================== TASK-P1-121 item 6: sideEffects =====================
' The per-call accumulators behind the additive `sideEffects` audit block.
' Design note (honest): P4-105 concluded there is no single RECEIPT BUILDER shared by
' every tool. There IS, however, a single place every tool's answer passes through:
' CallToolInner's text writer. SealedReceipt() therefore splices the block in there,
' which covers success AND failure receipts of every mutating tool in one edit instead
' of N per-tool edits - and cannot be forgotten by a future tool.
' ===========================================================================
Private Sub SeReset()
    Set m_seCreated = New Collection
    Set m_seRenamed = New Collection
    Set m_seRemoved = New Collection
    Set m_seFiles = New Collection
    Set Util.g_seBackups = New Collection
    Set m_seUndo = New Collection
End Sub

Private Sub SeAdd(ByVal coll As Object, ByVal value As String)
    On Error Resume Next
    If coll Is Nothing Then Exit Sub
    If Len(value) = 0 Then Exit Sub
    coll.Add value
    Err.Clear
End Sub

' Recorded by the writers. Public so the sibling modules (VbFmt) can take part.
Public Sub SeNoteCreated(ByVal name As String)
    SeAdd m_seCreated, name
End Sub

Public Sub SeNoteRemoved(ByVal name As String)
    SeAdd m_seRemoved, name
End Sub

Public Sub SeNoteRenamed(ByVal fromName As String, ByVal toName As String)
    On Error Resume Next
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("from") = fromName
    r("to") = toName
    If Not m_seRenamed Is Nothing Then m_seRenamed.Add r
    Err.Clear
End Sub

Public Sub SeNoteFile(ByVal path As String)
    SeAdd m_seFiles, path
End Sub

' (the .bak sink is Util.SeNoteBackup / Util.g_seBackups - shared with VbFmt.bas)

Private Function SideEffectsDict() As Object
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim fw As Object
    Set fw = New Collection
    Dim i As Long
    If m_seCreated Is Nothing Then SeReset
    If Not m_dirWrites Is Nothing Then
        For i = 1 To m_dirWrites.Count
            ManifestAdd fw, CStr(m_dirWrites(i))
        Next
    End If
    If Not m_seFiles Is Nothing Then
        For i = 1 To m_seFiles.Count
            ManifestAdd fw, CStr(m_seFiles(i))
        Next
    End If
    If m_seCreated Is Nothing Then Set m_seCreated = New Collection
    If m_seRenamed Is Nothing Then Set m_seRenamed = New Collection
    If m_seRemoved Is Nothing Then Set m_seRemoved = New Collection
    If Util.g_seBackups Is Nothing Then Set Util.g_seBackups = New Collection
    If m_seUndo Is Nothing Then Set m_seUndo = New Collection
    Set d("componentsCreated") = m_seCreated
    Set d("componentsRenamed") = m_seRenamed
    Set d("componentsRemoved") = m_seRemoved
    Set d("filesWritten") = fw
    Set d("backupsCreated") = Util.g_seBackups
    Set d("undoSnapshots") = m_seUndo
    Set SideEffectsDict = d
    Err.Clear
End Function

' Exact field names (recorded for the contract): componentsCreated[] -
' componentsRenamed[{from,to}] - componentsRemoved[] - filesWritten[] -
' backupsCreated[] - undoSnapshots[]. Additive only.
Private Sub SideEffectsInto(ByVal root As Object)
    On Error Resume Next
    If root Is Nothing Then Exit Sub
    Set root("sideEffects") = SideEffectsDict()
    Err.Clear
End Sub

Private Function SideEffectsJson() As String
    On Error Resume Next
    SideEffectsJson = JsonStringify(SideEffectsDict())
    Err.Clear
End Function

' Splice the block into a mutating tool's answer text. Gated three ways so it can never
' corrupt a non-object answer: (1) the tool must be state-changing, (2) the text must be
' a single JSON object, (3) the text must not already carry the key.
Private Function SealedReceipt(ByVal toolName As String, ByVal body As String) As String
    SealedReceipt = body
    On Error Resume Next
    If BodyIsNothing(body) Then Exit Function
    If Len(body) < 2 Then Exit Function
    If Right$(body, 1) <> "}" Then Exit Function
    If Left$(LTrim$(body), 1) <> "{" Then Exit Function
    If IsMutatingTool(toolName) Then
        If InStr(1, body, Chr$(34) & "sideEffects" & Chr$(34), vbTextCompare) > 0 Then Exit Function
        SealedReceipt = Left$(body, Len(body) - 1) & "," & Chr$(34) & "sideEffects" & Chr$(34) & ":" & SideEffectsJson() & "}"
    Else
        ' TASK-P1-126 (MCP-08/09): EVERY read-only receipt carries the in-flight block,
        ' so a caller can ask who is in flight, since when, and whether it is a read or
        ' a write from a tool that the busy gate never refuses. Reading a non-object
        ' answer is impossible here: the three guards above already returned.
        If InStr(1, body, Chr$(34) & "inflight" & Chr$(34), vbTextCompare) > 0 Then Exit Function
        m_inflightStamped = m_inflightStamped + 1
        SealedReceipt = Left$(body, Len(body) - 1) & "," & Chr$(34) & "inflight" & Chr$(34) & ":" & InflightJson() & "}"
    End If
    Err.Clear
End Function

Private Function BodyIsNothing(ByVal body As String) As Boolean
    BodyIsNothing = (Len(body) = 0)
End Function

Private Function ToolExportComponent(ByVal name As String, ByVal path As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim target As String
    If Len(path) = 0 Then
        target = comp.FileNames(1)
        If Len(target) = 0 Then Err.Raise vbObjectError + 49, "CodeTools", "Component has no file name; pass an explicit path"
    Else
        target = ResolvePath(path)
    End If
    ' Export writes a COPY; it must NOT reassign the component's file
    ' name (SaveAs would point the project at the export path and
    ' break later compiles).
    Dim hasDes As Boolean
    hasDes = False
    On Error Resume Next
    hasDes = Not (comp.Designer Is Nothing)
    If Err.Number <> 0 Then
        Err.Clear
        hasDes = False
    End If
    On Error GoTo 0
    Dim desNote As String
    desNote = ""
    If hasDes Then
        ' A designer component carries a binary layout section that the code
        ' text cannot represent. VBComponent.Export is the cleanest path but is
        ' absent from some VB6 VBE builds (error 438, measured here), so try it,
        ' then copy the component's saved project files, then fall back to text.
        Dim desOk As Boolean
        desOk = False
        On Error Resume Next
        comp.Export target
        If Err.Number = 0 Then
            desOk = True
        Else
            Err.Clear
        End If
        On Error GoTo 0
        Dim frx As String
        If Len(target) > 4 Then
            frx = Left$(target, Len(target) - 4) & ".frx"
        Else
            frx = target & ".frx"
        End If
        If desOk Then
            Dim rootD As Object
            Set rootD = CreateObject("Scripting.Dictionary")
            rootD("component") = comp.Name
            rootD("file") = target
            If CtlFileExists(frx) Then rootD("frx") = frx
            rootD("saved") = True
            rootD("note") = "Designer component exported whole via VBComponent.Export (.frm/.frx); the project's own file assignment is unchanged."
            ToolExportComponent = JsonStringify(rootD)
            Exit Function
        End If
        Dim src As String
        src = ""
        On Error Resume Next
        ' FileNames is a parameterised property (FileNames(index) As String), so
        ' asking for .Count is a compile error; read index 1 and swallow any
        ' failure (unsaved component).
        src = comp.FileNames(1)
        If Err.Number <> 0 Then
            Err.Clear
            src = ""
        End If
        On Error GoTo 0
        If Len(src) > 0 Then
            If CtlFileExists(src) Then
                Dim fso As Object
                Set fso = CreateObject("Scripting.FileSystemObject")
                fso.CopyFile src, target, True
                Dim srcFrx As String
                srcFrx = ""
                If Len(src) > 4 Then srcFrx = Left$(src, Len(src) - 4) & ".frx"
                If Len(srcFrx) > 0 Then
                    If CtlFileExists(srcFrx) Then fso.CopyFile srcFrx, frx, True
                End If
                Dim rootC As Object
                Set rootC = CreateObject("Scripting.Dictionary")
                rootC("component") = comp.Name
                rootC("file") = target
                If CtlFileExists(frx) Then rootC("frx") = frx
                rootC("saved") = True
                rootC("copiedFrom") = src
                rootC("note") = "VBComponent.Export is unavailable in this VBE build, so the component's saved project file was copied (it reflects the last saved state, not unsaved edits); the project's own file assignment is unchanged."
                ToolExportComponent = JsonStringify(rootC)
                Exit Function
            End If
        End If
        desNote = "Designer component: only the code section was written (VBComponent.Export unavailable and the component has no saved file yet, so the layout section is missing)."
    End If
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    Dim text As String
    If Not cm Is Nothing Then
        If cm.CountOfLines > 0 Then text = cm.Lines(1, cm.CountOfLines)
    End If
    Dim fnum As Integer
    fnum = FreeFile
    Open target For Output As #fnum
    Print #fnum, text;
    Close #fnum
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("file") = target
    root("saved") = True
    If Len(desNote) > 0 Then
        root("note") = desNote
    Else
        root("note") = "Code-only component exported as text (same bytes as the code module)."
    End If
    ToolExportComponent = JsonStringify(root)
End Function

Private Function ToolRenameComponent(ByVal name As String, ByVal newName As String) As String
    If Len(newName) = 0 Then Err.Raise vbObjectError + 50, "CodeTools", "new_name is required"
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    comp.Name = newName
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("oldName") = name
    root("newName") = comp.Name
    SeNoteRenamed name, comp.Name
    ToolRenameComponent = JsonStringify(root)
End Function

Private Function ToolReloadComponent(ByVal name As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    ' TASK-P1-119 / D-1115-02: read the component's identity BEFORE Reload.
    ' Measured (TASK-P1-115 J arm, re-measured TASK-P1-119): comp.Name read AFTER Reload reports
    ' VB6's DEFAULT component name (a standard module => "Module1") whenever that component's code
    ' pane has never been opened in this IDE session, and the true name once it has. FindComponent
    ' has just matched LCase$(comp.Name) against the requested name, so the value read HERE is
    ' provably the real name; reading it after the reload is what made the receipt lie.
    Dim compName As String
    compName = comp.Name
    ' TASK-P1-119 / D-1115-01: Reload re-reads the file the component is BOUND to (FileNames(1)),
    ' which is not necessarily the file the caller edited (import_component mode=file binds a
    ' project-local COPY). Report that bound path so the caller can decide whether the reload
    ' could have covered the edit at all.
    Dim boundFile As String
    boundFile = ""
    On Error Resume Next
    boundFile = comp.FileNames(1)
    If Err.Number <> 0 Then
        Err.Clear
        boundFile = ""
    End If
    On Error GoTo 0
    comp.Reload
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("reloaded") = compName
    root("requested") = name
    root("file") = boundFile
    ToolReloadComponent = JsonStringify(root)
End Function

' =====================================================================
' code module / code pane tools
' =====================================================================

Private Function ToolAddCode(ByVal component As String, ByVal code As String, ByVal file As String, ByVal validateMode As String) As String
    ' TASK-P1-059(d): segmented timing, the same shape replace_code/edit_by_anchor report.
    Dim tEntryA As Long
    tEntryA = GetTickCount()
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(component)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
    Dim addedFrom As String
    Dim body As String
    Dim srcPath As String
    srcPath = ""
    If Len(file) > 0 Then
        srcPath = ResolvePath(file)
        ' TASK-P1-121 item 2 (S2/G-3): append the FILE'S BODY, not its designer head.
        ' P4-101/103 measured that AddFromFile injects the header verbatim into the pane
        ' AND renames the target component from the file's Attribute VB_Name. The
        ' structural strip removes the head; verifyPane=False because the object being
        ' split here is a source FILE, not a code pane (spec D-178).
        body = StrippedDiskText(Build.ReadLogAnsi(srcPath), "", False)
        addedFrom = "file"
    ElseIf Len(code) > 0 Then
        body = code
        addedFrom = "string"
    Else
        Err.Raise vbObjectError + 52, "CodeTools", "code or file is required"
    End If
    ' TASK-P1-126 (MCP-06/MCP-04): the same PRE-WRITE structural gate as replace_code.
    Dim pairRefA As String
    pairRefA = PreWritePairingRefusal("add_code", comp.Name, body, "the appended block (addedFrom=" & addedFrom & ")")
    If Len(pairRefA) > 0 Then
        ToolAddCode = pairRefA
        Exit Function
    End If
    Dim cur As String
    ' TASK-P1-062: guard the empty-component case (CountOfLines = 0 => Lines(1,0) = error 5);
    ' add_code into an EMPTY module is legal and must write the first line.
    If cm.CountOfLines > 0 Then cur = cm.Lines(1, cm.CountOfLines)
    Dim curLines() As String
    curLines = TextLines(cur)
    Dim tReadA As Long
    tReadA = GetTickCount()
    ' ---- TASK-P1-020 beta-2: appending must not extend a continued statement ----
    If IsContinuationTail(curLines, cm.CountOfLines + 1) Then
        ToolAddCode = GuardReject(comp.Name, "add_code", "inside_continuation", _
            "the last line ends with a VB continuation _, so appending would extend that logical statement. Close the statement first, then add the code.", _
            0, 0, cm.CountOfLines)
        Exit Function
    End If
    ' ---- TASK-P1-020 beta-3: lint the RESULT before it is written (warn-only) ----
    Dim predicted As String
    predicted = cur
    If Len(predicted) > 0 And Len(body) > 0 Then predicted = predicted & vbCrLf
    predicted = predicted & Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf)
    Dim pack As Object
    ' TASK-P1-054: region target = the first line the appended code lands on.
    Dim tPredictA As Long
    tPredictA = GetTickCount()
    Set pack = LintPackForEdit(predicted, KindOfName(component), cm.CountOfLines, LineCount(cur) + 1)
    Dim tLintA As Long
    tLintA = GetTickCount()
    ' TASK-P1-024 gamma-4
    Dim acref As String
    acref = StrictRefusal(component, "add_code", pack, validateMode)
    If Len(acref) > 0 Then
        ToolAddCode = acref
        Exit Function
    End If
    Dim tWriteA As Long
    tWriteA = GetTickCount()
    ' TASK-P1-121 item 2 (S2/G-3): insert the head-stripped BODY at the END of the module.
    ' AddFromFile (the old call) appended at the end - that is what the tool description
    ' promises ("Append code to a component's code module"). VB6's CodeModule.AddFromString
    ' inserts at the CARET, NOT at the end: measured in this session, a freshly loaded
    ' 3-line module grew at LINE 2 when AddFromString was used, i.e. the naive translation
    ' of this item would have been a regression (the code would land next to the caret
    ' instead of at the end, and the caret is wherever the IDE left it).
    ' InsertLines(CountOfLines + 1) is the explicit end-append and also covers an EMPTY
    ' module (Line = 1). The `code` STRING branch is deliberately left exactly as it was -
    ' it used AddFromString before this card and still does; its caret dependence is
    ' pre-existing and is reported as a new finding, not silently changed here.
    If Len(srcPath) > 0 Then
        cm.InsertLines cm.CountOfLines + 1, body
    Else
        cm.AddFromString code
    End If
    Dim tWriteEndA As Long
    tWriteEndA = GetTickCount()
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("addedFrom") = addedFrom
    Dim paneAfter As String
    paneAfter = CompText(comp.Name)
    Dim textTotal As Long, ideTotal As Long
    textTotal = LineCount(paneAfter)
    ideTotal = cm.CountOfLines
    root("newTotalLines") = ideTotal
    root("ideTotalLines") = ideTotal
    root("textTotalLines") = textTotal
    If ideTotal <> textTotal Then root("syncWarning") = SyncWarningText(ideTotal, textTotal)
    Set root("requestedRange") = RangeDict("add_code", cm.CountOfLines + 1, cm.CountOfLines + 1)
    Set root("appliedRange") = AppliedRangeDict(cm.CountOfLines + 1, cm.CountOfLines)
    root("clamped") = False
    Set root("coords") = CoordsDictOfName(comp.Name, cm.CountOfLines + 1, "ide")
    AddPaneDiskFields root, comp.Name, paneAfter
    CopyLintPack root, pack
    root("readMs") = tReadA - tEntryA
    root("predictMs") = tPredictA - tReadA
    root("lintMs") = tLintA - tPredictA
    root("writeMs") = tWriteEndA - tWriteA
    ' TASK-P1-093: PURE ADDITION. insMs is the AddFromString/AddFromFile interval (the
    ' same interval writeMs covers; writeMs itself is unchanged) and add_code never
    ' deletes, so delMs is 0. Key shapes/values of the pre-existing keys do not move.
    root("insMs") = tWriteEndA - tWriteA
    root("delMs") = 0
    root("writeCalls") = 1
    ' TASK-P1-075: AddFromString/AddFromFile appends one block => hunks = 1 (same meaning
    ' as everywhere else in the write receipts; CONTRACTS §3).
    root("hunks") = 1
    root("totalMs") = GetTickCount() - tEntryA
    ToolAddCode = ReceiptJson(root)
End Function

Private Function ToolCreateEventProc(ByVal component As String, ByVal objName As String, ByVal eventName As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(component)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
    Dim line As Long
    line = cm.CreateEventProc(eventName, objName)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("proc") = objName & "_" & eventName
    root("startLine") = line
    ToolCreateEventProc = JsonStringify(root)
End Function

Private Function ToolFindInComponent(ByVal component As String, ByVal target As String, ByVal startLine As Long, ByVal startCol As Long, ByVal endLine As Long, ByVal endCol As Long, ByVal wholeWord As Boolean, ByVal matchCase As Boolean, ByVal patternSearch As Boolean) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(component)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
    If startLine < 1 Then startLine = 1
    If startCol < 1 Then startCol = 1
    If endLine < 1 Then endLine = cm.CountOfLines
    If endCol < 1 Then endCol = 255
    Dim found As Boolean
    found = cm.Find(target, startLine, startCol, endLine, endCol, wholeWord, matchCase, patternSearch)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("found") = found
    root("startLine") = startLine
    root("startColumn") = startCol
    root("endLine") = endLine
    root("endColumn") = endCol
    ' TASK-P1-020 alpha: cm.Find is ByRef on the line/column arguments, so startLine
    ' is the FOUND line (an IDE CodeModule line) once found; coords states that.
    Set root("coords") = CoordsDictOfName(comp.Name, startLine, "ide")
    ToolFindInComponent = JsonStringify(root)
End Function

Private Function ToolGetCodePane(ByVal component As String) As String
    Dim cp As VBIDE.CodePane
    Set cp = GetCodePane(component)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = cp.CodeModule.Name
    root("topLine") = cp.TopLine
    root("countOfVisibleLines") = cp.CountOfVisibleLines
    root("view") = CodePaneViewName(cp.CodePaneView)
    Dim sl As Long, sc As Long, el As Long, ec As Long
    cp.GetSelection sl, sc, el, ec
    root("selectionStartLine") = sl
    root("selectionStartColumn") = sc
    root("selectionEndLine") = el
    root("selectionEndColumn") = ec
    On Error Resume Next
    root("windowCaption") = cp.Window.Caption
    On Error GoTo 0
    ToolGetCodePane = JsonStringify(root)
End Function

Private Function ToolSetCodePane(ByVal component As String, ByVal selection As String, ByVal topLine As Long, ByVal view As String, ByVal activate As Boolean) As String
    Dim cp As VBIDE.CodePane
    Set cp = GetCodePane(component)
    ' TASK-P1-164 (R18-2): set_codepane moves the SELECTION and the SCROLL position, never
    ' the text. The pane text is captured here and asserted at the end.
    Dim pAssertComp As String, pAssertBefore As String
    Dim pAssertOn As Boolean
    pAssertComp = cp.CodeModule.Name
    pAssertBefore = PaneTextOfName(pAssertComp, pAssertOn)
    Dim applied As Object
    Set applied = New Collection
    If Len(selection) > 0 Then
        Dim parts() As String
        parts = Split(selection, ",")
        If UBound(parts) < 3 Then Err.Raise vbObjectError + 53, "CodeTools", "selection must be startLine,startCol,endLine,endCol"
        cp.SetSelection CLng(Val(parts(0))), CLng(Val(parts(1))), CLng(Val(parts(2))), CLng(Val(parts(3)))
        applied.Add "selection"
    End If
    If topLine > 0 Then
        cp.TopLine = topLine
        applied.Add "topLine"
    End If
    Dim ignoredView As Boolean
    ignoredView = False
    If Len(view) > 0 Then
        ' TASK-P1-114: \`view\` is no longer DECLARED in this tool's inputSchema. VB6EXT v5.3
        ' exposes CodePane.CodePaneView as GET-ONLY (its put-map holds only TopLine) and has
        ' no \`View\` property at all; a real-IDE scan of all 34 command bars found no view
        ' toggle. The request is still refused by reality - but it is now REPORTED
        ' (ignoredArgs) instead of being answered with a silent no-op (discipline 38).
        applied.Add "view-ignored"
        ignoredView = True
    End If
    If activate Then
        cp.Show
        applied.Add "activated"
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = cp.CodeModule.Name
    Set root("applied") = applied
    If ignoredView Then
        root("ignoredArgs") = Array("view")
        root("ignoredArgsNote") = "set_codepane has no view-mode parameter: VB6's extensibility model exposes " & _
            "CodePane.CodePaneView as READ-ONLY (VB6EXT v5.3, the _CodePane put-map holds only TopLine) and the IDE's " & _
            "34 command bars carry no view toggle, so the code pane view cannot be switched from an add-in. The rest of " & _
            "this call WAS applied (see applied) - the argument is reported here rather than silently ignored. Inspect the " & _
            "current mode with get_codepane.view; switch it with the IDE's own Full Module / Procedure View buttons."
    End If
    ' R18-2: a selection/scroll change must not rewrite the pane text.
    PaneAssertFinish root, pAssertComp, pAssertBefore, pAssertOn, "set_codepane"
    ToolSetCodePane = JsonStringify(root)
End Function

Private Function ToolGetSelection() As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Dim ap As VBIDE.VBProject
    Set ap = m_vbe.ActiveVBProject
    If Not ap Is Nothing Then root("activeProject") = ap.Name
    Dim svc As VBIDE.VBComponent
    Set svc = m_vbe.SelectedVBComponent
    If Not svc Is Nothing Then
        root("selectedComponent") = svc.Name
        root("selectedComponentType") = ComponentTypeName(svc.Type)
    End If
    Dim aw As VBIDE.Window
    Set aw = m_vbe.ActiveWindow
    If Not aw Is Nothing Then root("activeWindow") = aw.Caption
    Dim cp As VBIDE.CodePane
    Set cp = m_vbe.ActiveCodePane
    If cp Is Nothing Then
        Dim fp As VBIDE.VBProject
        Set fp = GetProject()
        Dim fc As VBIDE.VBComponent
        For Each fc In fp.VBComponents
            Set cp = fc.CodeModule.CodePane
            If Not cp Is Nothing Then Exit For
        Next
    End If
    If Not cp Is Nothing Then
        root("activeCodeModule") = cp.CodeModule.Name
        Dim sl As Long, sc As Long, el As Long, ec As Long
        cp.GetSelection sl, sc, el, ec
        root("cursorLine") = sl
        root("cursorColumn") = sc
        root("selectionEndLine") = el
        root("selectionEndColumn") = ec
    End If
    On Error GoTo 0
    ToolGetSelection = JsonStringify(root)
End Function

Private Function ToolActivateComponent(ByVal name As String, ByVal window As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Select Case LCase$(window)
        Case "code"
            Dim cm As VBIDE.CodeModule
            Set cm = comp.CodeModule
            If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & name
            cm.CodePane.Show
        Case "designer"
            comp.DesignerWindow
        Case Else
            comp.Activate
    End Select
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("window") = window
    ToolActivateComponent = JsonStringify(root)
End Function

' =====================================================================
' file tools (disk)
' =====================================================================

Private Function ToolListFiles(ByVal path As String, ByVal pattern As String, ByVal recursive As Boolean, ByVal maxEntries As Long) As String
    If maxEntries <= 0 Then maxEntries = 500
    Dim base As String
    base = ProjectDir()
    Dim target As String
    If Len(path) = 0 Then
        target = base
        If Len(target) = 0 Then Err.Raise vbObjectError + 30, "CodeTools", "No project open and no path given"
    Else
        target = ResolvePath(path)
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(target) Then Err.Raise vbObjectError + 31, "CodeTools", "Directory not found: " & target
    Dim patterns As Variant
    If Len(pattern) = 0 Then pattern = "*"
    patterns = Split(pattern, ";")
    Dim results As Object
    Set results = New Collection
    Dim count As Long
    Dim truncated As Boolean
    WalkDir fso, fso.GetFolder(target), patterns, results, base, recursive, maxEntries, count, truncated
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("path") = target
    root("count") = count
    root("truncated") = truncated
    Set root("files") = results
    ToolListFiles = JsonStringify(root)
End Function

Private Function ToolReadFile(ByVal path As String, ByVal maxLines As Long) As String
    Dim full As String
    full = ResolvePath(path)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(full) Then Err.Raise vbObjectError + 32, "CodeTools", "File not found: " & full
    Dim content As String
    content = ReadTextAuto(full)
    If maxLines > 0 Then
        Dim lines() As String
        lines = Split(Replace(content, vbCrLf, vbLf), vbLf)
        If UBound(lines) >= maxLines Then
            Dim head() As String
            ReDim head(0 To maxLines - 1)
            Dim i As Long
            For i = 0 To maxLines - 1
                MaybePump
                head(i) = lines(i)
            Next
            content = Join(head, vbCrLf) & vbCrLf & "... (truncated at " & CStr(maxLines) & " lines; total " & CStr(UBound(lines) + 1) & ")"
        End If
    End If
    ToolReadFile = content
End Function

Private Function ToolWriteFile(ByVal path As String, ByVal content As String, ByVal encoding As String, ByVal createDirs As Boolean, _
                              ByVal backup As Boolean, ByVal validateMode As String) As String
    Dim full As String
    full = ResolvePath(path)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If createDirs Then
        Dim parent As String
        parent = fso.GetParentFolderName(full)
        EnsureDir parent
    End If
    Dim enc As String
    enc = LCase$(encoding)
    If enc = "auto" Then
        If IsVbFileExt(full) Then enc = "ansi" Else enc = "utf-8"
    End If
    Dim backupPath As String
    backupPath = ""
    If backup And fso.FileExists(full) Then
        backupPath = full & ".bak"
        If fso.FileExists(backupPath) Then fso.DeleteFile backupPath, True
        fso.CopyFile full, backupPath
    End If
    Dim bytes() As Byte
    Select Case enc
        Case "utf-8"
            bytes = Utf8Encode(content)
        Case "utf-16"
            bytes = Utf16Encode(content)
        Case Else
            bytes = AnsiEncode(content)
            enc = "ansi"
    End Select
    ' TASK-P1-020 beta-3: lint a VB source file BEFORE it is written (warn-only).
    Dim lintPk As Object
    If IsVbFileExt(full) Then
        Set lintPk = LintPack(content, InferKind(full))
        ' TASK-P1-024 gamma-4: refuse BEFORE the file is touched
        Dim wfref As String
        wfref = StrictRefusal(full, "write_file", lintPk, validateMode)
        If Len(wfref) > 0 Then
            ToolWriteFile = wfref
            Exit Function
        End If
    End If
    ' TASK-P1-021: VB6's "Open ... For Binary Access Write" does NOT truncate the
    ' file, so writing SHORTER content left the previous tail behind (measured: a
    ' 50-byte file written down to 10 bytes kept 40 bytes of the old content and
    ' reported bytesWritten=10). Delete the target first so the bytes on disk are
    ' exactly the new content. The backup above has already been taken.
    On Error Resume Next
    If fso.FileExists(full) Then fso.DeleteFile full, True
    Err.Clear
    On Error GoTo 0
    Dim fnum As Integer
    fnum = FreeFile
    Open full For Binary Access Write As #fnum
    If (UBound(bytes) - LBound(bytes)) >= 0 Then Put #fnum, , bytes
    Close #fnum
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("file") = full
    root("encoding") = enc
    root("bytesWritten") = UBound(bytes) - LBound(bytes) + 1
    If Len(backupPath) > 0 Then root("backup") = backupPath
    If IsProjectFile(full) Then
        root("note") = "File belongs to the open project. Use reload_component to load it into the IDE, or use replace_code for live in-memory edits."
    End If
    CopyLintPack root, lintPk
    ToolWriteFile = JsonStringify(root)
End Function

Private Function ToolSearchInFiles(ByVal query As String, ByVal path As String, ByVal pattern As String, ByVal isRegex As Boolean, ByVal caseSensitive As Boolean, ByVal maxResults As Long) As String
    If maxResults <= 0 Then maxResults = 200
    Dim base As String
    base = ProjectDir()
    Dim target As String
    If Len(path) = 0 Then
        target = base
        If Len(target) = 0 Then Err.Raise vbObjectError + 30, "CodeTools", "No project open and no path given"
    Else
        target = ResolvePath(path)
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(target) Then Err.Raise vbObjectError + 31, "CodeTools", "Directory not found: " & target
    Dim patterns As Variant
    If Len(pattern) = 0 Then pattern = DEFAULT_SEARCH_PATTERNS
    patterns = Split(pattern, ";")
    Dim re As Object
    Set re = Nothing
    If isRegex Then
        Set re = CreateObject("VBScript.RegExp")
        re.Pattern = query
        re.Global = False
        re.IgnoreCase = Not caseSensitive
    End If
    Dim results As Object
    Set results = New Collection
    Dim count As Long
    Dim truncated As Boolean
    Dim filesScanned As Long
    SearchWalk fso, fso.GetFolder(target), patterns, re, query, isRegex, caseSensitive, maxResults, results, count, truncated, filesScanned, base
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("query") = query
    root("regex") = isRegex
    root("caseSensitive") = caseSensitive
    root("path") = target
    root("filesScanned") = filesScanned
    root("resultsCount") = count
    root("truncated") = truncated
    Set root("results") = results
    ToolSearchInFiles = JsonStringify(root)
End Function

' =====================================================================
' build / run tools
' =====================================================================

Private Function ToolSaveProject(Optional ByVal path As String) As String
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    ' Hard guard: VB6 raises an access violation in USER32 (WER 0xc0000005,
    ' measured 2026-09-21) while saving a project that references a component
    ' file which is missing from disk. Refuse instead of taking the IDE down.
    Dim missing As String
    missing = MissingComponentFiles(p)
    ' VB6 silently DROPS a component whose file is missing, so the dangling
    ' reference is only visible in the .vbp text. Saving while it is there
    ' crashes the IDE (access violation in USER32), so refuse and say exactly
    ' which line to fix.
    Dim dangling As String
    dangling = ""
    If Len(p.FileName) > 0 Then dangling = MissingVbpFiles(p.FileName)
    If Len(missing) > 0 Then dangling = dangling & missing
    If Len(dangling) > 0 Then
        OpLog "save_project REFUSED: dangling references " & dangling
        Err.Raise vbObjectError + 88, "CodeTools", "refusing to save: the project references component files that do not exist on disk, and saving this state crashes VB6 (access violation in USER32). Restore the files or remove the dangling lines from the .vbp first. Missing: " & dangling
    End If
    AuditAdd "save_project", "begin", p.Name & " gitpath=" & p.FileName
    Dim pPath As String
    pPath = Trim$(path)
    ' TASK-P1-077: remember every file this call writes, so the receipt can list them.
    Dim vbpWritten As String
    vbpWritten = ""
    On Error Resume Next
    If Len(pPath) > 0 Then
        If Len(p.FileName) = 0 Or LCase$(p.FileName) <> LCase$(pPath) Then
            Err.Clear
            p.SaveAs pPath
            If Err.Number = 0 Then vbpWritten = pPath Else Err.Clear
        End If
    End If
    On Error GoTo 0
    Dim baseDir As String
    baseDir = ProjectDir()
    Dim saved As Object
    Set saved = New Collection
    Dim skipped As Object
    Set skipped = New Collection
    Dim cWrites As Object
    Set cWrites = New Collection
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        On Error Resume Next
        Dim dirty As Boolean
        dirty = comp.IsDirty
        On Error GoTo 0
        ' forms: IsDirty is unreliable (designer edits do not always
        ' set it) - treat components that were never saved (no file
        ' name) as needing a save so run/make never trigger VB6's
        ' modal Save dialogs
        If dirty Or Len(comp.FileNames(1)) = 0 Then
            If Len(comp.FileNames(1)) > 0 Then
                Dim cf As String
                cf = comp.FileNames(1)
                comp.SaveAs cf
                ManifestAdd cWrites, cf
                saved.Add comp.Name
            ElseIf Len(baseDir) > 0 Then
                Dim ext As String
                Select Case comp.Type
                    Case 1  ' vbext_ct_StdModule
                        ext = ".bas"
                    Case 2  ' vbext_ct_ClassModule
                        ext = ".cls"
                    Case 3, 5, 6  ' MSForm / VBForm / VBMDIForm
                        ' (this build reports standard forms as type 5)
                        ext = ".frm"
                    Case 4  ' vbext_ct_ResFile
                        ext = ".res"
                    Case Else
                        ext = ""
                End Select
                If Len(ext) > 0 Then
                    comp.SaveAs baseDir & "\" & comp.Name & ext
                    ManifestAdd cWrites, baseDir & "\" & comp.Name & ext
                    saved.Add comp.Name
                Else
                    skipped.Add comp.Name
                End If
            Else
                skipped.Add comp.Name
            End If
        End If
    Next
    Dim projectSaved As Boolean
    projectSaved = False
    On Error Resume Next
    If Not p.Saved Then
        If Len(p.FileName) > 0 Then
            Err.Clear
            p.SaveAs p.FileName
            If Err.Number = 0 Then vbpWritten = p.FileName Else Err.Clear
            projectSaved = True
        End If
    Else
        projectSaved = True
    End If
    On Error GoTo 0
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("project") = p.Name
    root("projectSaved") = projectSaved
    Set root("savedComponents") = saved
    Set root("skippedComponents") = skipped
    ' TASK-P1-077: the write manifest is ALWAYS present (empty array when the call wrote
    ' nothing, e.g. an already-clean project). It lists every file this call wrote: the project
    ' file when it was (re)saved plus each component file written by SaveAs. That is a superset
    ' of "files inside the project directory" - a component bound to a file outside it is listed
    ' too, so nothing can hide behind the field name.
    Dim pdw As Object
    Set pdw = New Collection
    ManifestAdd pdw, vbpWritten
    Dim wi As Long
    On Error Resume Next
    For wi = 1 To cWrites.Count
        ManifestAdd pdw, cWrites(wi)
    Next
    On Error GoTo 0
    ManifestMergeUndo pdw
    Set root("projectDirWrites") = pdw
    ToolSaveProject = JsonStringify(root)
End Function

' Resolve the real build target and never leave it empty.
' VB6 opens its Make Project file dialog whenever BuildFileName is empty,
' and a title-based dialog match cannot rescue that.
Private Function MakeTargetPath(ByVal p As VBIDE.VBProject, ByVal outputPath As String) As String
    Dim dir As String, stem As String, ext As String
    If Len(outputPath) > 0 Then
        MakeTargetPath = ResolvePath(outputPath)
        Exit Function
    End If
    If Len(p.BuildFileName) > 0 Then
        MakeTargetPath = p.BuildFileName
        Exit Function
    End If
    If Len(p.FileName) = 0 Then Exit Function
    dir = Left$(p.FileName, InStrRev(p.FileName, "\"))
    stem = Mid$(p.FileName, Len(dir) + 1)
    If LCase$(Right$(stem, 4)) = ".vbp" Then stem = Left$(stem, Len(stem) - 4)
    Select Case p.Type
        Case 3: ext = ".dll"
        Case 4: ext = ".ocx"
        Case Else: ext = ".exe"
    End Select
    MakeTargetPath = dir & stem & ext
End Function

' TASK-P1-125 (MCP-13 / MCP-02 / MCP-03): the artifact is a PRODUCT, not scratch
' space. The old implementation DELETED the existing output before every Make
' ("delete the target first, or VB6 pops an overwrite dialog"), so a Make that
' failed, timed out, or merely put up VB6's own Make Project box left the caller
' with NO artifact at all - a real add-in EXE was measured going 5,750,837 B ->
' ENOENT / size 0, and had to restore it from a backup. Now the old artifact is
' MOVED ASIDE (never deleted), restored on every failure path, and its
' {exists,size,mtime,sha256} is reported before and after as output{}.
' =====================================================================
' TASK-P1-125 (MCP-13 / MCP-02 / MCP-03 / MCP-28): artifact bookkeeping.
' =====================================================================

' {path,exists,size,mtime,mtimeText,sha256} of a build artifact. A build receipt
' can then never be handed back without saying what happened to the product itself.
Private Function ArtifactFullSnap(ByVal path As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = path
    d("exists") = False
    d("size") = 0
    d("mtime") = 0
    d("mtimeText") = ""
    ' TASK-P1-126 (F1): the digest key is OMITTED (and a reason given) whenever the
    ' digest cannot be computed. Before this, an exclusive-locked artifact was reported
    ' as sha256=E3B0C442... - the digest of an EMPTY input - while exists=true and
    ' size=16384. Publishing a wrong value is worse than omitting the field.
    d("sha256Available") = False
    d("sha_unavailable_reason") = "artifact_does_not_exist"
    On Error Resume Next
    If Len(path) > 0 Then
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        If fso.FileExists(path) Then
            d("exists") = True
            d("size") = CDbl(fso.GetFile(path).Size)
            d("mtime") = CDbl(fso.GetFile(path).DateLastModified)
            d("mtimeText") = Format$(fso.GetFile(path).DateLastModified, "yyyy-mm-dd hh:nn:ss")
            Dim shaWhy As String
            Dim hv As String
            shaWhy = ""
            hv = Build.FileSha256Why(path, shaWhy)
            If Len(hv) > 0 Then
                d("sha256") = hv
                d("sha256Available") = True
                d("sha_unavailable_reason") = ""
            Else
                d("sha_unavailable_reason") = shaWhy
            End If
        End If
    End If
    On Error GoTo 0
    Set ArtifactFullSnap = d
End Function

' TASK-P1-127 (R1): reading an ABSENT dictionary key with the default-property syntax
' auto-ADDS it as Empty (a Scripting.Dictionary gotcha) and CStr(Empty) is "", so a
' snapshot whose digest could not be computed must be read deliberately, not with CStr.
Private Function SnapSha(ByVal snap As Object) As String
    SnapSha = ""
    If snap Is Nothing Then Exit Function
    Dim v As Variant
    On Error Resume Next
    Err.Clear
    v = snap("sha256")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If IsNull(v) Then Exit Function
    SnapSha = CStr(v)
End Function

' TASK-P1-125 (MCP-28 / S point 5): is the artifact actually LOCKED by someone?
' WMI names a lock owner only when the file is a running process image, so an
' independent test is needed: try to take an exclusive lock ourselves. Opening
' For Binary with Lock Read Write never truncates the file, and it fails with
' error 70 when another process holds the file.
Private Function ArtifactWriteLocked(ByVal path As String) As Boolean
    ArtifactWriteLocked = False
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    If Dir$(path) = "" Then Exit Function
    Dim fn As Long
    fn = FreeFile
    Err.Clear
    Open path For Binary Access Read Lock Read Write As #fn
    If Err.Number <> 0 Then
        ArtifactWriteLocked = True
        Err.Clear
        Exit Function
    End If
    Close #fn
    On Error GoTo 0
End Function

' TASK-P1-125 (MCP-28): which process holds the artifact? Windows has no API that
' names a file lock owner, so the RUNNING IMAGE is matched instead: a client that
' holds the built .exe has it as its process image. Best effort - 0 means unknown,
' and a caller must never present 0 as "not locked".
Private Function ArtifactHolderPid(ByVal path As String) As Long
    ArtifactHolderPid = 0
    If Len(path) = 0 Then Exit Function
    On Error Resume Next
    If Dir$(path) = "" Then Exit Function
    ' TASK-P1-126 (F3): ExecutablePath alone MISSED a real holder (the verifier measured
    ' the holder as pid 19340 while this function answered 0). Three passes now:
    '   (1) same normalized full path, (2) same FILE NAME, (3) the artifact path named in
    '   a process command line. A pass-2/3 hit is weaker, so it is only used when no
    '   exact hit exists - and a genuinely unnamed holder still answers 0, which the
    '   receipt must describe as UNKNOWN (never as "unlocked").
    Dim svc As Object, col As Object, o As Object, exe As String, cmd As String
    Dim wantNorm As String, wantBase As String
    wantNorm = LCase$(Replace(path, "/", "\"))
    wantBase = LCase$(Mid$(path, InStrRev(path, "\") + 1))
    Dim weakPid As Long
    weakPid = 0
    Set svc = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    If svc Is Nothing Then Exit Function
    Set col = svc.ExecQuery("SELECT ProcessId, ExecutablePath, CommandLine FROM Win32_Process")
    If col Is Nothing Then Exit Function
    For Each o In col
        exe = ""
        exe = CStr(o.ExecutablePath)
        If Len(exe) > 0 Then
            If StrComp(LCase$(Replace(exe, "/", "\")), wantNorm, vbTextCompare) = 0 Then
                ArtifactHolderPid = CLng(o.ProcessId)
                Exit For
            End If
            If weakPid = 0 Then
                If LCase$(Mid$(exe, InStrRev(exe, "\") + 1)) = wantBase Then weakPid = CLng(o.ProcessId)
            End If
        End If
        If weakPid = 0 Then
            cmd = ""
            cmd = CStr(o.CommandLine)
            If Len(cmd) > 0 Then
                If InStr(1, cmd, path, vbTextCompare) > 0 Then weakPid = CLng(o.ProcessId)
            End If
        End If
    Next
    If ArtifactHolderPid = 0 Then ArtifactHolderPid = weakPid
    On Error GoTo 0
End Function

' TASK-P1-125 (discipline 98): ONE key set for make_project on EVERY path, so no
' branch can silently drop a field the caller depends on. Callers overwrite what
' they know; everything else keeps the documented default.
' =====================================================================
' TASK-P1-132 (A10): LAA (IMAGE_FILE_LARGE_ADDRESS_AWARE) is applied by the
' MCP add-in ITSELF - no external EditBin.exe is required.
' VB6's linker does not set PE Characteristics bit 0x0020, and the author's
' release script used to patch it afterwards, so every build that went through
' make_project silently lost the flag. The PE header is now rewritten in place
' (exactly 2 bytes, one Put), re-read to VERIFY, and rolled back when the
' verification fails. An artifact that already carries the flag is untouched.
' =====================================================================
' TASK-P1-132 (A10): is the IMAGE signed?  IMAGE_DIRECTORY_ENTRY_SECURITY (index 4)
' of the optional header's data directory names the certificate table. A non-zero
' size means an Authenticode signature is attached, and rewriting the PE header
' afterwards would invalidate it - so LaaPatchJson refuses instead of breaking it
' silently. Only the SIZE field is read (the offset differs between PE32/PE32+).
Private Function PeCertTableSize(ByVal path As String, ByRef certBytes As Long, ByRef ok As Boolean, ByRef errMsg As String) As Boolean
    PeCertTableSize = False
    certBytes = 0
    ok = False
    errMsg = ""
    Dim f As Integer
    On Error GoTo Bad
    f = FreeFile
    Open path For Binary Access Read As #f
    Dim sig As Long
    Get #f, 1, sig
    If (sig And &HFFFF&) <> &H5A4D& Then
        Close #f
        errMsg = "not a PE image"
        Exit Function
    End If
    Dim lfanew As Long
    Get #f, 61, lfanew
    If lfanew <= 0 Then
        Close #f
        errMsg = "e_lfanew is invalid"
        Exit Function
    End If
    Dim optPos As Long
    optPos = lfanew + 25
    Dim magic As Integer
    Get #f, optPos, magic
    Dim sizePos As Long
    If magic = &H10B Then
        sizePos = optPos + 128 + 4
    ElseIf magic = &H20B Then
        sizePos = optPos + 144 + 4
    Else
        Close #f
        errMsg = "unknown optional-header magic " & CStr(magic)
        Exit Function
    End If
    Dim sv As Long
    Get #f, sizePos, sv
    Close #f
    certBytes = sv
    ok = True
    PeCertTableSize = True
    Exit Function
Bad:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    On Error Resume Next
    Close #f
    errMsg = "could not read the optional header: VB error " & CStr(en) & " (" & ed & ")"
    Err.Clear
End Function

' TASK-P1-132 (A10 constraint 2): LAA only enlarges a 32-bit image's user address
' space on a 64-bit Windows (or a 32-bit Windows booted with /3GB). Report whether
' the flag can actually take effect on THIS machine instead of implying it always does.
Private Function LaaMachineEffective() As Boolean
    Dim pa As String, pa32 As String
    pa = UCase$(Environ$("PROCESSOR_ARCHITECTURE"))
    pa32 = UCase$(Environ$("PROCESSOR_ARCHITEW6432"))
    LaaMachineEffective = (Len(pa32) > 0) Or (pa = "AMD64") Or (pa = "IA64")
End Function

' =====================================================================
' TASK-P1-132 (A10): WHERE does the LAA decision come from?
' Priority: per-call parameter laa > the .vbp NAMED section [VB6MCP] (LAA / LAANote)
' > the optional sidecar <vbpDir>\.vb6mcp.json > the environment variable VB6MCP_LAA
' > the global default (ON).
' Survival rule for the .vbp source: a NAMED section that is present WHEN THE PROJECT
' IS LOADED is re-emitted verbatim by VB6 on every save and survives long-term; a
' section APPENDED externally WHILE the project is open is overwritten by that save -
' so write the .vbp with the IDE CLOSED, or use the sidecar file, which can be written
' at any time. It is the ROOT-area keys that the in-memory model overwrites (the
' CondComp case is that class), NOT a named section.
' =====================================================================
Private Function IsBoolLike(ByVal v As Variant, ByRef outBool As Boolean) As Boolean
    IsBoolLike = False
    outBool = False
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        outBool = CBool(v)
        IsBoolLike = True
        Exit Function
    End If
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    If s = "true" Or s = "1" Or s = "yes" Or s = "on" Then
        outBool = True
        IsBoolLike = True
        Exit Function
    End If
    If s = "false" Or s = "0" Or s = "no" Or s = "off" Then
        outBool = False
        IsBoolLike = True
    End If
End Function

' TASK-P1-132 (A10): the PRIMARY switch lives in the project's OWN .vbp, in a
' custom section:
'     [VB6MCP]
'     LAA=1
'     LAANote=this project needs >2GB
' EXACT survival rule (measured): a section that was PRESENT WHEN THE PROJECT WAS
' LOADED survives every IDE save, because VB6 re-emits it verbatim; a section
' APPENDED to the .vbp from outside WHILE the project is open is LOST on the IDE's
' next save, because it was never in the in-memory project model. So write this
' section while the IDE has the project CLOSED, or use the sidecar file below,
' which can be written at any time.
' Keys are read case-insensitively; 1/true/yes/on = true, 0/false/no/off = false;
' an ILLEGAL value is reported, never silently defaulted.
' Returns True when the section produced a decision; False with an EMPTY errMsg
' when the section/key is simply absent (the caller falls through to the next
' source); False with a NON-EMPTY errMsg when the section exists but is unusable.
Private Function LaaVbpSection(ByVal vbp As String, ByRef found As Boolean, ByRef keyPresent As Boolean, _
        ByRef want As Boolean, ByRef note As String, ByRef errMsg As String) As Boolean
    LaaVbpSection = False
    found = False
    keyPresent = False
    want = False
    note = ""
    errMsg = ""
    Dim txt As String
    txt = Build.ReadLogAnsi(vbp)
    If Len(txt) = 0 Then Exit Function
    Dim lines() As String
    lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim i As Long, ln As String, low As String
    Dim inSec As Boolean, rawVal As String, keyName As String
    Dim seenLaa As Boolean, seenNote As String
    inSec = False
    seenLaa = False
    seenNote = ""
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        low = LCase$(ln)
        If Len(ln) > 0 Then
            If Left$(ln, 1) = "[" Then
                ' a new section ends the one we were reading
                If inSec Then Exit For
                If low = "[vb6mcp]" Then
                    inSec = True
                    found = True
                End If
            ElseIf inSec Then
                Dim eq As Long
                eq = InStr(ln, "=")
                If eq > 0 Then
                    keyName = LCase$(Trim$(Left$(ln, eq - 1)))
                    rawVal = Trim$(Mid$(ln, eq + 1))
                    If keyName = "laa" Then
                        Dim vb As Boolean
                        If IsBoolLike(rawVal, vb) Then
                            want = vb
                            seenLaa = True
                        Else
                            errMsg = "the [VB6MCP] section of " & vbp & " has LAA=" & rawVal & ", which is not a recognised boolean (1/true/yes/on or 0/false/no/off), so this project's LAA setting is UNKNOWN. Nothing was silently defaulted: the LAA step was SKIPPED. Fix the value or remove the key (a per-call laa argument overrides it)."
                            LaaVbpSection = False
                            Exit Function
                        End If
                    ElseIf keyName = "laanote" Then
                        seenNote = rawVal
                    End If
                End If
            End If
        End If
    Next
    keyPresent = seenLaa
    note = seenNote
    If seenLaa Then
        LaaVbpSection = True
    ElseIf found Then
        ' the section exists without a LAA key: not a decision, fall through
        LaaVbpSection = False
    End If
End Function

Private Function LaaSidecarPath(ByVal vbp As String) As String
    Dim fso As Object
    Dim dir As String
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    dir = fso.GetParentFolderName(vbp)
    On Error GoTo 0
    If Len(dir) = 0 Then Exit Function
    LaaSidecarPath = PathJoin(dir, ".vb6mcp.json")
End Function

' True  = the decision is usable (want/src/projFile/note are set)
' False = a configuration source EXISTS but is unusable; errMsg says what, and
'         NO default is substituted (the LAA step must be skipped, not guessed).
Private Function LaaResolve(ByVal vbp As String, ByVal laaGiven As Boolean, ByVal laaParam As Boolean, _
        ByRef want As Boolean, ByRef src As String, ByRef projFile As String, _
        ByRef note As String, ByRef errMsg As String) As Boolean
    LaaResolve = True
    want = True
    src = "global-default"
    note = ""
    errMsg = ""
    projFile = LaaSidecarPath(vbp)
    If laaGiven Then
        want = laaParam
        src = "param"
        note = "the laa argument was passed on THIS call, so it decides"
        Exit Function
    End If
    ' (2) the project's OWN [VB6MCP] section (PRIMARY source; VB6 preserves
    ' unknown sections when it saves the .vbp).
    Dim vFound As Boolean, vKey As Boolean, vWant As Boolean, vNote As String, vErr As String
    vFound = False
    vKey = False
    vWant = False
    vNote = ""
    vErr = ""
    If LaaVbpSection(vbp, vFound, vKey, vWant, vNote, vErr) Then
        want = vWant
        src = "vb-project"
        note = "the [VB6MCP] section of the .vbp decides for this project"
        If Len(vNote) > 0 Then note = note & ": " & vNote
        Exit Function
    End If
    If Len(vErr) > 0 Then
        LaaResolve = False
        src = "vb-project-error"
        errMsg = vErr
        Exit Function
    End If
    ' (3) the optional sidecar <vbpDir>\.vb6mcp.json overrides for projects whose
    ' .vbp is managed by external tools.
    If Len(projFile) > 0 Then
        Dim has As Boolean
        has = False
        On Error Resume Next
        Err.Clear
        has = (Dir$(projFile) <> "")
        If Err.Number <> 0 Then has = False
        Err.Clear
        On Error GoTo 0
        If has Then
            Dim txt As String
            txt = Build.ReadLogAnsi(projFile)
            Dim parsed As Variant
            If Not JsonLib.JsonParse(txt, parsed) Then
                LaaResolve = False
                src = "project-file-error"
                errMsg = "the project sidecar " & projFile & " exists but is NOT valid JSON, so this project's LAA setting is UNKNOWN. Nothing was silently defaulted: the LAA step was SKIPPED. Fix the file (e.g. {""laa"": true}) or delete it - a per-call laa argument overrides it."
                Exit Function
            End If
            If Not IsObject(parsed) Then
                LaaResolve = False
                src = "project-file-error"
                errMsg = "the project sidecar " & projFile & " is valid JSON but not a JSON OBJECT, so this project's LAA setting is UNKNOWN. The LAA step was SKIPPED (nothing was defaulted)."
                Exit Function
            End If
            Dim o As Object
            Set o = parsed
            If Not o.Exists("laa") Then
                LaaResolve = False
                src = "project-file-error"
                errMsg = "the project sidecar " & projFile & " has no laa key, so this project's LAA setting is UNKNOWN. The LAA step was SKIPPED (nothing was defaulted); add {""laa"": true|false} or delete the file."
                Exit Function
            End If
            Dim pv As Variant
            Dim pb As Boolean
            pv = o("laa")
            If IsBoolLike(pv, pb) Then
                want = pb
                src = "project-file"
                note = "the project sidecar decides for this project"
                If o.Exists("laaNote") Then note = note & ": " & CStr(o("laaNote"))
                Exit Function
            End If
            LaaResolve = False
            src = "project-file-error"
            errMsg = "the project sidecar " & projFile & " has a laa value that is not true/false, so this project's LAA setting is UNKNOWN. The LAA step was SKIPPED (nothing was defaulted)."
            Exit Function
        End If
    End If
    Dim ev As String
    ev = Trim$(Environ$("VB6MCP_LAA"))
    If Len(ev) > 0 Then
        Dim eb As Boolean
        If IsBoolLike(ev, eb) Then
            want = eb
            src = "env"
            note = "the environment variable VB6MCP_LAA decides"
            Exit Function
        End If
        LaaResolve = False
        src = "env-error"
        errMsg = "the environment variable VB6MCP_LAA is set to a value that is not true/false, so the LAA setting is UNKNOWN. The LAA step was SKIPPED (nothing was defaulted)."
        Exit Function
    End If
    note = "no per-call laa argument, no project sidecar and no VB6MCP_LAA: the global default (ON) applies"
End Function

' A configuration refusal must LOOK like one: never a green receipt.
Private Sub LaaConfigRefusal(ByVal root As Object, ByVal projFile As String, ByVal errMsg As String)
    root("laaRequested") = False
    root("laaApplicable") = False
    root("laaBefore") = False
    root("laaAfter") = False
    root("laaSet") = False
    root("laaAlready") = False
    root("laaBytesPatched") = 0
    root("laaError") = errMsg
    root("laaNote") = "the LAA configuration could not be resolved, so NO PE header was touched. This is NOT a silent fallback to the default."
    root("laaOk") = False
    root("laaFlag") = "IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020)"
    root("laaMethod") = "the add-in rewrites the PE header itself; an external EditBin.exe is NOT required"
    root("laaProjectFile") = projFile
    root("laaMachineEffective") = LaaMachineEffective()
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    Set root("laa") = e
End Sub

Private Function PeLaaRead(ByVal path As String, ByRef chars As Long, ByRef charsOff As Long, ByRef errMsg As String) As Boolean
    PeLaaRead = False
    chars = 0
    charsOff = 0
    errMsg = ""
    Dim f As Integer
    Dim nFileBytes As Double
    On Error GoTo Bad
    nFileBytes = FileLen(path)
    f = FreeFile
    Open path For Binary Access Read As #f
    Dim sig As Long
    Get #f, 1, sig
    If (sig And &HFFFF&) <> &H5A4D& Then
        Close #f
        errMsg = "not a PE image: the DOS signature at offset 0 is not MZ"
        Exit Function
    End If
    Dim lfanew As Long
    Get #f, 61, lfanew
    If lfanew <= 0 Then
        Close #f
        errMsg = "not a PE image: e_lfanew reads " & CStr(lfanew)
        Exit Function
    End If
    If CDbl(lfanew) + 24 > nFileBytes Then
        Close #f
        errMsg = "truncated PE header: e_lfanew=" & CStr(lfanew) & " but the file is only " & CStr(CLng(nFileBytes)) & " bytes"
        Exit Function
    End If
    Dim peSig As Long
    Get #f, lfanew + 1, peSig
    If peSig <> &H4550& Then
        Close #f
        errMsg = "not a PE image: no PE signature at e_lfanew"
        Exit Function
    End If
    Dim lo As Byte, hi As Byte
    charsOff = lfanew + 23
    Get #f, charsOff, lo
    Get #f, charsOff + 1, hi
    Close #f
    chars = CLng(lo) + CLng(hi) * 256
    PeLaaRead = True
    Exit Function
Bad:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    On Error Resume Next
    Close #f
    errMsg = "could not read the PE header: VB error " & CStr(en) & " (" & ed & ")"
    Err.Clear
End Function

Private Function PeWriteChars(ByVal path As String, ByVal charsOff As Long, ByVal newChars As Long, ByRef errMsg As String) As Boolean
    PeWriteChars = False
    errMsg = ""
    Dim f As Integer
    On Error GoTo Bad
    f = FreeFile
    ' Binary + Access Read Write opens WITHOUT truncation: only the 2 bytes that
    ' are Put change, and one Put of a 2-byte array is a single write.
    Open path For Binary Access Read Write As #f
    Dim two(0 To 1) As Byte
    two(0) = CByte(newChars And &HFF&)
    two(1) = CByte((newChars And &HFF00&) \ &H100&)
    Put #f, charsOff, two
    Close #f
    PeWriteChars = True
    Exit Function
Bad:
    Dim en As Long, ed As String
    en = Err.Number
    ed = Err.Description
    On Error Resume Next
    Close #f
    errMsg = "VB error " & CStr(en) & " (" & ed & ")"
    If en = 70 Then errMsg = errMsg & " - the artifact is read-only or held open by another process"
    If en = 75 Then errMsg = errMsg & " - the artifact path is not writable"
    Err.Clear
End Function

' Returns {applicable, laaBefore, laaAfter, laaSet, already, bytesPatched, peOk, error, note}
Private Function LaaPatchJson(ByVal artifact As String, ByVal want As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("artifact") = artifact
    d("requested") = want
    d("applicable") = False
    d("laaBefore") = False
    d("laaAfter") = False
    d("laaSet") = False
    d("already") = False
    d("bytesPatched") = 0
    d("peOk") = False
    d("error") = ""
    d("note") = ""
    d("signedImage") = False
    d("certificateTableBytes") = 0
    d("laaSkippedReason") = ""
    d("laaRefusedReason") = ""
    If Not want Then
        d("note") = "laa=false: the caller switched the LAA step off for this call; the artifact was not examined"
        Set LaaPatchJson = d
        Exit Function
    End If
    If Len(artifact) = 0 Then
        d("error") = "laa: the artifact path is unknown, so the PE header cannot be examined or patched"
        Set LaaPatchJson = d
        Exit Function
    End If
    Dim exists As Boolean
    exists = False
    On Error Resume Next
    Err.Clear
    exists = (Dir$(artifact) <> "")
    If Err.Number <> 0 Then exists = False
    Err.Clear
    On Error GoTo 0
    If Not exists Then
        d("error") = "laa: the artifact does not exist on disk yet (" & artifact & "), so there is no PE header to patch"
        Set LaaPatchJson = d
        Exit Function
    End If
    ' ---- TASK-P1-132 (A10 constraint 5): NEVER patch this add-in's own DLL ---------
    ' An ActiveX DLL's LAA is decided by the HOST process's address space, so setting
    ' it on VB6MCP.dll would be meaningless. make_project/build patch only the TARGET
    ' product's artifact; if the target IS this DLL, skip it and say so.
    Dim ownPath As String
    ownPath = ""
    On Error Resume Next
    Err.Clear
    ownPath = BuildPath()
    On Error GoTo 0
    If Len(ownPath) > 0 Then
        If PathsSame(ownPath, artifact) Then
            d("laaSkippedReason") = "own_addin_dll"
            d("note") = "the target artifact IS this add-in's own DLL (the ActiveX server being built); LAA on an ActiveX DLL is decided by the host process, so the flag was deliberately NOT written"
            Set LaaPatchJson = d
            Exit Function
        End If
    End If
    Dim chars As Long, charsOff As Long, peErr As String
    Dim peOk As Boolean
    peOk = PeLaaRead(artifact, chars, charsOff, peErr)
    d("peOk") = peOk
    If Not peOk Then
        d("error") = "laa: " & peErr
        Set LaaPatchJson = d
        Exit Function
    End If
    d("applicable") = True
    Dim isLaa As Boolean
    isLaa = ((chars And &H20&) <> 0)
    d("laaBefore") = isLaa
    d("laaAfter") = isLaa
    If isLaa Then
        d("already") = True
        d("note") = "idempotent: the artifact already carries IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020); nothing was written"
        Set LaaPatchJson = d
        Exit Function
    End If
    ' ---- TASK-P1-132 (A10 constraint 1): code signing must come AFTER the flag -----
    ' An Authenticode certificate lives in a PE data directory that a header edit
    ' invalidates. A signed artifact is REFUSED rather than silently broken; the flag
    ' must be set BEFORE signing (signing is the last step), or laa=false must be
    ' passed deliberately. Nothing is written on this path.
    Dim certBytes As Long, certOk As Boolean, certErr As String
    certBytes = 0
    certOk = False
    certErr = ""
    If PeCertTableSize(artifact, certBytes, certOk, certErr) Then
        If certBytes <> 0 Then
            d("signedImage") = True
            d("certificateTableBytes") = certBytes
            d("laaRefusedReason") = "signed_image"
            d("error") = "laa: the artifact carries an Authenticode certificate table (" & CStr(certBytes) & " bytes). Rewriting the PE header would INVALIDATE that signature, so the flag was NOT written. Set the flag BEFORE code signing (signing must be the last step), or pass laa=false to accept a non-LAA artifact deliberately."
            Set LaaPatchJson = d
            Exit Function
        End If
    Else
        d("certCheckError") = certErr
    End If
    Dim wrErr As String
    If Not PeWriteChars(artifact, charsOff, chars Or &H20&, wrErr) Then
        d("error") = "laa: could not write the PE characteristics word: " & wrErr
        Set LaaPatchJson = d
        Exit Function
    End If
    Dim chars2 As Long, off2 As Long, peErr2 As String
    If PeLaaRead(artifact, chars2, off2, peErr2) Then
        If (chars2 And &H20&) <> 0 Then
            d("laaAfter") = True
            d("laaSet") = True
            d("bytesPatched") = 2
            d("note") = "IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) was written into the PE header by the MCP add-in (no EditBin.exe) and the file was re-read to verify"
            Set LaaPatchJson = d
            Exit Function
        End If
    End If
    Dim rbErr As String
    Dim rbOk As Boolean
    rbOk = PeWriteChars(artifact, charsOff, chars, rbErr)
    d("laaAfter") = False
    d("laaSet") = False
    If rbOk Then
        d("error") = "laa: the write was not confirmed by a re-read; the original characteristics word was put back (rollback ok). The artifact is unchanged."
    Else
        d("error") = "laa: the write was not confirmed by a re-read AND the rollback failed (" & rbErr & "). The PE header may be damaged - rebuild, or restore from the .before-build copy."
    End If
    Set LaaPatchJson = d
End Function

' Publishes the LAA result into a receipt under BOTH the flat names the card
' asks for (laaBefore/laaAfter/laaSet) and a laa{} object with the detail.
Private Sub ApplyLaaInto(ByVal root As Object, ByVal artifact As String, ByVal want As Boolean)
    Dim r As Object
    Set r = LaaPatchJson(artifact, want)
    root("laaRequested") = want
    root("laaApplicable") = CBool(r("applicable"))
    root("laaBefore") = CBool(r("laaBefore"))
    root("laaAfter") = CBool(r("laaAfter"))
    root("laaSet") = CBool(r("laaSet"))
    root("laaAlready") = CBool(r("already"))
    root("laaBytesPatched") = CLng(r("bytesPatched"))
    root("laaError") = CStr(r("error"))
    root("laaNote") = CStr(r("note"))
    root("laaOk") = (Len(CStr(r("error"))) = 0)
    root("laaFlag") = "IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020)"
    root("laaMethod") = "the add-in rewrites the PE header itself; an external EditBin.exe is NOT required"
    root("laaSkippedReason") = CStr(r("laaSkippedReason"))
    root("laaRefusedReason") = CStr(r("laaRefusedReason"))
    root("laaSignedImage") = CBool(r("signedImage"))
    root("laaCertificateTableBytes") = CLng(r("certificateTableBytes"))
    root("laaMachineEffective") = LaaMachineEffective()
    root("laaMachineEffectiveNote") = ""
    If Not CBool(root("laaMachineEffective")) Then root("laaMachineEffectiveNote") = "LAA only enlarges a 32-bit image's user address space on 64-bit Windows (or 32-bit Windows booted with /3GB); on this machine the flag is a promise that does not change the address space"
    Set root("laa") = r
End Sub

' =====================================================================
' TASK-P1-132 (MCP-lead-2): DIAGNOSTIC MODE - an isolated VB6.EXE /make on a
' COPY of the .vbp plus its components. The in-IDE Make route has no compiler
' subprocess, and running /make on the project the IDE holds yields exit 0 and
' no log; a copy has neither problem. The real project and the real artifact
' are never touched.
' =====================================================================
Private Function PathJoin(ByVal dir As String, ByVal name As String) As String
    If Len(dir) = 0 Then
        PathJoin = name
        Exit Function
    End If
    If Right$(dir, 1) = Chr$(92) Then
        PathJoin = dir & name
    Else
        PathJoin = dir & Chr$(92) & name
    End If
End Function

Private Function IsProjectComponentKey(ByVal low As String) As Boolean
    IsProjectComponentKey = False
    If Left$(low, 5) = "form=" Then IsProjectComponentKey = True
    If Left$(low, 7) = "module=" Then IsProjectComponentKey = True
    If Left$(low, 6) = "class=" Then IsProjectComponentKey = True
    If Left$(low, 12) = "usercontrol=" Then IsProjectComponentKey = True
    If Left$(low, 13) = "propertypage=" Then IsProjectComponentKey = True
    If Left$(low, 9) = "designer=" Then IsProjectComponentKey = True
    If Left$(low, 11) = "relateddoc=" Then IsProjectComponentKey = True
    If Left$(low, 10) = "resfile32=" Then IsProjectComponentKey = True
End Function

' Copies every file entry of the .vbp into tmpDir (relative paths preserved),
' rewrites ExeName32 so the isolated make writes INSIDE tmpDir, and returns the
' path of the copied .vbp.
Private Function CopyProjectForIsolatedMake(ByVal vbp As String, ByVal tmpDir As String, _
        ByRef outVbp As String, ByRef nCopied As Long, ByRef missing As String, ByRef errMsg As String) As Boolean
    CopyProjectForIsolatedMake = False
    outVbp = ""
    nCopied = 0
    missing = ""
    errMsg = ""
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo Bad
    Dim text As String
    text = Build.ReadLogAnsi(vbp)
    If Len(text) = 0 Then
        errMsg = "the .vbp could not be read as text: " & vbp
        Exit Function
    End If
    Dim srcDir As String
    srcDir = fso.GetParentFolderName(vbp)
    Dim lines() As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim outLines() As String
    ReDim outLines(LBound(lines) To UBound(lines))
    Dim i As Long, ln As String, low As String, val As String
    Dim exeName As String, exeFound As Boolean
    exeName = ""
    exeFound = False
    For i = LBound(lines) To UBound(lines)
        ln = lines(i)
        low = LCase$(Trim$(ln))
        outLines(i) = ln
        If IsProjectComponentKey(low) Then
            Dim p As Long
            p = InStr(ln, "=")
            If p > 0 Then
                val = Trim$(Mid$(ln, p + 1))
                Dim sc As Long
                sc = InStr(val, ";")
                If sc > 0 Then val = Trim$(Mid$(val, sc + 1))
                If Len(val) > 0 Then
                    Dim src As String
                    src = PathJoin(srcDir, val)
                    If Not fso.FileExists(src) Then
                        missing = missing & val & ";"
                    Else
                        Dim dst As String
                        dst = PathJoin(tmpDir, val)
                        Dim dd As String
                        dd = fso.GetParentFolderName(dst)
                        On Error Resume Next
                        If Not fso.FolderExists(dd) Then fso.CreateFolder dd
                        On Error GoTo Bad
                        fso.CopyFile src, dst, True
                        nCopied = nCopied + 1
                        Dim ext As String
                        ext = LCase$(fso.GetExtensionName(val))
                        If ext = "frm" Or ext = "ctl" Or ext = "dsr" Then
                            Dim frx As String
                            frx = Left$(src, Len(src) - 4) & ".frx"
                            If fso.FileExists(frx) Then
                                fso.CopyFile frx, Left$(dst, Len(dst) - 4) & ".frx", True
                                nCopied = nCopied + 1
                            End If
                        End If
                    End If
                End If
            End If
        ElseIf Left$(low, 10) = "exename32=" Then
            exeName = Trim$(Mid$(ln, InStr(ln, "=") + 1))
            exeFound = True
        End If
    Next
    If Len(exeName) = 0 Then exeName = fso.GetFileName(vbp)
    Dim stem As String, extn As String
    stem = fso.GetBaseName(exeName)
    extn = fso.GetExtensionName(exeName)
    If Len(extn) = 0 Then extn = "exe"
    Dim isoName As String
    isoName = stem & "_isolated." & extn
    If exeFound Then
        For i = LBound(outLines) To UBound(outLines)
            If LCase$(Left$(Trim$(outLines(i)), 10)) = "exename32=" Then outLines(i) = "ExeName32=" & isoName
        Next
    Else
        ReDim Preserve outLines(LBound(outLines) To UBound(outLines) + 1)
        outLines(UBound(outLines)) = "ExeName32=" & isoName
    End If
    outVbp = PathJoin(tmpDir, fso.GetFileName(vbp))
    ProjShip.WriteTextFile outVbp, Join(outLines, vbCrLf)
    CopyProjectForIsolatedMake = True
    Exit Function
Bad:
    errMsg = "isolated copy failed: VB error " & CStr(Err.Number) & " (" & Err.Description & ")"
End Function

Private Sub DiagnoseInto(ByVal root As Object, ByVal vbp As String, ByVal timeoutMs As Long, ByVal reason As String)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("attempted") = True
    d("reason") = reason
    d("route") = "isolated_copy_make"
    d("ok") = False
    d("tmpDir") = ""
    d("vbpCopy") = ""
    d("copiedFiles") = 0
    d("missingFiles") = ""
    d("logPath") = ""
    d("log") = ""
    d("logBytes") = 0
    d("errorCount") = 0
    d("exitCode") = -1
    d("artifactPath") = ""
    d("artifactExists") = False
    d("artifactSize") = 0
    d("error") = ""
    d("note") = "compiles a COPY of the .vbp plus its component files in a temp directory with an isolated VB6.EXE /make; the project in the IDE and the real artifact are NOT touched. The copy reflects the files ON DISK, so unsaved pane edits are not in it."
    Dim fe As Object
    Set fe = CreateObject("Scripting.Dictionary")
    fe("file") = ""
    fe("line") = 0
    fe("lineSpace") = "disk"
    fe("text") = ""
    Set d("firstError") = fe
    On Error GoTo Bad
    Dim tmpBase As String
    tmpBase = Environ$("TEMP")
    If Len(tmpBase) = 0 Then tmpBase = Environ$("TMP")
    If Len(tmpBase) = 0 Then
        d("error") = "no TEMP directory is defined, so an isolated copy cannot be made"
        Set root("diagnosis") = d
        Exit Sub
    End If
    Dim tmpDir As String
    tmpDir = PathJoin(tmpBase, "vb6mcp_diag_" & Format$(Now, "yyyymmdd_hhnnss") & "_" & CStr(GetTickCount() And &HFFFF&))
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    fso.CreateFolder tmpDir
    d("tmpDir") = tmpDir
    Dim vbpCopy As String, nCopied As Long, missing As String, copyErr As String
    If Not CopyProjectForIsolatedMake(vbp, tmpDir, vbpCopy, nCopied, missing, copyErr) Then
        d("error") = copyErr
        Set root("diagnosis") = d
        Exit Sub
    End If
    d("vbpCopy") = vbpCopy
    d("copiedFiles") = nCopied
    d("missingFiles") = missing
    Dim logPath As String
    logPath = PathJoin(tmpDir, "isolated_make.log")
    d("logPath") = logPath
    ' ToolBuild already knows how to run an isolated VB6.EXE /make, read the log
    ' and locate the first error; reuse it. laaWant=False: the isolated image is
    ' throwaway - the REAL artifact is patched by the caller.
    Dim json As String
    json = ToolBuild(vbpCopy, "", logPath, timeoutMs, True, False, False)
    Dim parsed As Variant
    If Not JsonLib.JsonParse(json, parsed) Then
        d("error") = "the isolated /make receipt could not be parsed"
        Set root("diagnosis") = d
        Exit Sub
    End If
    If Not IsObject(parsed) Then
        d("error") = "the isolated /make receipt was not a JSON object"
        Set root("diagnosis") = d
        Exit Sub
    End If
    Dim o As Object
    Set o = parsed
    If o.Exists("exitCode") Then d("exitCode") = CLng(Val(CStr(o("exitCode"))))
    If o.Exists("logBytes") Then d("logBytes") = CLng(Val(CStr(o("logBytes"))))
    If o.Exists("log") Then d("log") = CStr(o("log"))
    If o.Exists("errorCount") Then d("errorCount") = CLng(Val(CStr(o("errorCount"))))
    If o.Exists("artifact") Then d("artifactPath") = CStr(o("artifact"))
    If o.Exists("output") Then
        If IsObject(o("output")) Then
            Dim ob As Object
            Set ob = o("output")
            If ob.Exists("exists") Then d("artifactExists") = CBool(ob("exists"))
            If ob.Exists("size") Then d("artifactSize") = CDbl(Val(CStr(ob("size"))))
        End If
    End If
    If o.Exists("firstError") Then
        If IsObject(o("firstError")) Then
            Dim ofe As Object
            Set ofe = o("firstError")
            Dim fe2 As Object
            Set fe2 = CreateObject("Scripting.Dictionary")
            fe2("file") = CStr(ofe("file"))
            fe2("line") = CLng(Val(CStr(ofe("line"))))
            fe2("lineSpace") = "disk"
            fe2("text") = CStr(ofe("text"))
            Set d("firstError") = fe2
        End If
    End If
    If CLng(d("errorCount")) = 0 And CBool(d("artifactExists")) Then
        d("ok") = True
    ElseIf CLng(d("errorCount")) > 0 Then
        d("error") = "the isolated /make reported " & CStr(d("errorCount")) & " error(s); see firstError and log"
    Else
        d("error") = "the isolated /make produced no artifact and no error count; read log and see diagnosisUnavailableReason"
    End If
    Set root("diagnosis") = d
    Exit Sub
Bad:
    d("error") = "diagnosis failed: VB error " & CStr(Err.Number) & " (" & Err.Description & ")"
    On Error Resume Next
    Set root("diagnosis") = d
End Sub

Private Function MakeReceiptEmpty() As Object
    Dim m As Object
    Set m = CreateObject("Scripting.Dictionary")
    m("ok") = False
    m("compiled") = False
    m("compiled_means") = "compiled=true means ONLY that VBProject.MakeCompiledFile returned without raising. It is NOT evidence that anything was built or rewritten - judge the build by artifactChanged, verdict and output{exists,size,sha256}."
    m("verdict") = ""
    m("buildStarted") = False
    m("modalConfirmAttempted") = False
    m("usedMenuFallback") = False
    m("artifactPath") = ""
    m("artifactChanged") = False
    m("artifact_restored") = False
    m("artifact_backup_path") = ""
    m("artifact_locked") = False
    m("artifact_holder_pid") = 0
    m("modal_blocked") = False
    m("modal_stage") = ""
    m("host_modal_risk") = "none"
    m("host_modal_risk_reason") = ""
    m("host_modal_risk_class") = ""
    m("modalDialogs") = "[]"
    m("logPath") = ""
    m("log") = ""
    m("logAvailable") = False
    m("logUnavailableReason") = ""
    m("errorCount") = 0
    Dim fe As Object
    Set fe = CreateObject("Scripting.Dictionary")
    fe("file") = ""
    fe("line") = 0
    fe("lineSpace") = "disk"
    fe("text") = ""
    Set m("firstError") = fe
    m("firstErrorSource") = ""
    m("pending") = False
    m("pending_how_to_confirm") = ""
    m("route") = "ide"
    m("requestedRoute") = ""
    m("suggested_action") = ""
    m("recovery_hint") = ""
    m("note") = ""
    ' TASK-P1-125 (S point 1): the shape is ALIGNED with the build tool, which
    ' already reported file/reportedLine/diskLine/logPath - make_project must not
    ' invent a second vocabulary for the same facts.
    m("file") = ""
    m("reportedLine") = 0
    m("diskLine") = 0
    m("logBytes") = 0
    m("logTruncated") = False
    ' TASK-P1-125 (S point 5): artifact_locker_pid is the name the A-side asked for;
    ' 0 means UNKNOWN, never "not locked".
    m("artifact_locker_pid") = 0
    m("artifact_locked_advice") = ""
    ' TASK-P1-125 (S point 3): a failed in-IDE Make used to leave one more pair of
    ' boxes up per attempt (A-side measured count growing to 7 over 5 failures).
    m("dialogs_before") = 0
    m("dialogs_after") = 0
    m("dialogs_leaked") = 0
    m("dialogs_close_attempts") = 0
    m("dialogs_note") = ""
    ' TASK-P1-125 (MCP-01 / 2.3): the armed one-shot confirm that can reach INSIDE
    ' VB6's own modal loop while MakeCompiledFile is blocked in it.
    m("make_watchdog_armed") = False
    m("make_watchdog_clicked") = False
    m("make_watchdog_status") = ""
    ' ---- TASK-P1-132: LAA (A10) + the isolated diagnosis (MCP-lead-2), on EVERY path
    m("laaRequested") = True
    m("laaApplicable") = False
    m("laaBefore") = False
    m("laaAfter") = False
    m("laaSet") = False
    m("laaAlready") = False
    m("laaBytesPatched") = 0
    m("laaError") = ""
    m("laaNote") = ""
    m("laaOk") = False
    m("laaFlag") = "IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020)"
    m("laaMethod") = "the add-in rewrites the PE header itself; an external EditBin.exe is NOT required"
    m("laaMachineEffective") = False
    m("laaMachineEffectiveNote") = ""
    m("laaSignedImage") = False
    m("laaCertificateTableBytes") = 0
    m("laaSkippedReason") = ""
    m("laaRefusedReason") = ""
    m("laaSource") = "global-default"
    m("laaProjectFile") = ""
    m("laaSourceNote") = ""
    m("laaConfigError") = ""
    Dim laaEmpty As Object
    Set laaEmpty = CreateObject("Scripting.Dictionary")
    Set m("laa") = laaEmpty
    m("unbuiltReasonUnknown") = False
    m("diagnosisRoute") = ""
    m("diagnosisTmpDir") = ""
    m("diagnosisCopiedFiles") = 0
    m("diagnosisMissingFiles") = ""
    m("diagnosisError") = ""
    m("diagnosisNote") = ""
    Dim diagEmpty As Object
    Set diagEmpty = CreateObject("Scripting.Dictionary")
    diagEmpty("attempted") = False
    Set m("diagnosis") = diagEmpty
    Set m("output") = ArtifactFullSnap("")
    Set MakeReceiptEmpty = m
End Function

Private Function ToolMakeProject(ByVal outputPath As String, ByVal route As String, _
                                 ByVal outPath As String, ByVal timeoutMs As Long, _
                                 ByVal allowHostModal As Boolean, _
                                 ByVal laaGiven As Boolean, ByVal laaParam As Boolean, _
                                 ByVal diagnoseWant As Boolean) As String
    Dim t0 As Long
    t0 = GetTickCount()
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    ' MakeCompiledFile / the Make menu command open a modal Save dialog for
    ' unsaved projects. Save silently first, or fail with guidance.
    If Len(p.FileName) = 0 Then
        Err.Raise vbObjectError + 40, "CodeTools", "Project is unsaved: compiling would open VB6's Save dialog. Call save_project with a path first, then retry."
    End If
    ' ---- TASK-P1-125 (3.3): route=cli is a FIRST-CLASS route, not a footnote ---
    ' The user reported that VB6.EXE /make <vbp> /out <log> is the reliable path, and
    ' it is the only route that produces a compiler log. route=cli delegates to it
    ' (the build tool) and returns that receipt verbatim, so route/path say which
    ' route actually ran instead of pretending.
    Dim effRoute As String
    effRoute = LCase$(Trim$(route))
    If Len(m_makeRouteOverride) > 0 Then effRoute = m_makeRouteOverride
    If Len(effRoute) = 0 Then effRoute = "auto"
    If effRoute = "cli" Then
        If timeoutMs <= 0 Then timeoutMs = 180000
        Dim prevOverride As String
        prevOverride = m_makeRouteOverride
        ' break the cycle: if build decides the IDE holds this project it must fall
        ' back to the IN-IDE path, never back into route=cli.
        m_makeRouteOverride = "ide"
        Dim cliJson As String
        Dim cliErr As Long, cliDesc As String
        cliErr = 0
        cliDesc = ""
        On Error Resume Next
        cliJson = ToolBuild(p.FileName, "", outPath, timeoutMs, laaGiven, laaParam, False)
        cliErr = Err.Number
        cliDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        m_makeRouteOverride = prevOverride
        If cliErr <> 0 Then
            Dim er As Object
            Set er = MakeReceiptEmpty()
            er("ok") = False
            er("verdict") = "cli_route_failed"
            er("requestedRoute") = "cli"
            er("route") = "cli"
            er("logUnavailableReason") = "the isolated VB6.EXE /make subprocess could not be started: " & cliDesc
            er("recovery_hint") = "retry with route=ide (the default) to compile inside the IDE, or call the build tool directly with an explicit out_path."
            er("note") = "route=cli failed before a compiler log existed (" & CStr(cliErr) & ")"
            ToolMakeProject = JsonStringify(er)
            Exit Function
        End If
        ToolMakeProject = cliJson
        Exit Function
    End If
    On Error Resume Next
    Dim saveErr As Long
    Dim saveDesc As String
    saveErr = 0
    saveDesc = ""
    If Not p.Saved Then
        p.SaveAs p.FileName
        saveErr = Err.Number
        saveDesc = Err.Description
    End If
    On Error GoTo 0
    If saveErr <> 0 Then Err.Raise vbObjectError + 40, "CodeTools", "Could not save the project before compiling: " & saveDesc
    ' ---- TASK-P1-132 (A10 sidecar): resolve WHERE the LAA decision comes from ----
    Dim laaWant As Boolean, laaSrc As String, laaProjFile As String, laaSrcNote As String, laaCfgErr As String
    Dim laaUsable As Boolean
    laaUsable = LaaResolve(p.FileName, laaGiven, laaParam, laaWant, laaSrc, laaProjFile, laaSrcNote, laaCfgErr)
    ' ---- TASK-P1-125: the artifact BEFORE anything is touched ------------------
    Dim targetExe As String
    targetExe = MakeTargetPath(p, outputPath)
    Dim before As Object
    Set before = ArtifactFullSnap(targetExe)
    Dim root As Object
    Set root = MakeReceiptEmpty()
    root("requestedRoute") = effRoute
    root("artifactPath") = targetExe
    root("route") = "ide"
    Set root("output") = before
    ' ---- TASK-P1-125 (MCP-01 interlock): a box that ALREADY existed -------------
    ' A host-modal box that predates this call was not opened by us, so confirming it
    ' would put the add-in inside somebody else's modal loop - the exact stuck state
    ' the user hit (every later write tool answered reason=busy until the IDE was
    ' killed). Refuse, and hand back the hwnd/title/controls needed to decide.
    Dim preRisk As String, preReason As String, preClass As String, preDetails As String
    Dim preNDlg As Long
    preNDlg = WinOps.HostModalRiskProbe("#32770", preRisk, preReason, preClass, preDetails)
    root("dialogs_before") = preNDlg
    If preNDlg > 0 And preRisk = "known" And Not allowHostModal Then
        root("modal_blocked") = True
        root("modal_stage") = "preexisting"
        root("host_modal_risk") = preRisk
        root("host_modal_risk_reason") = preReason
        root("host_modal_risk_class") = preClass
        root("modalDialogs") = preDetails
        root("verdict") = "modal_blocked"
        root("logUnavailableReason") = "no compile was attempted, so no compiler log exists"
        root("suggested_action") = "read the box with read_active_dialogs and close it with close_active_dialogs (dry_run=true previews it, including host_modal_risk); pass allow_host_modal=true only if you deliberately want THIS call to confirm it."
        root("recovery_hint") = "Nothing was started and the artifact was not touched (see output{}). If a previous call is still in flight, its busy receipt carries inflight_tool / inflight_since_ms; read_active_dialogs still answers because read-only tools are not gated."
        root("note") = "a host-modal box was already up before this call, so make_project neither compiled nor touched the artifact"
        ' TASK-P1-125 (discipline 98): this early exit must still report the same
        ' shape, including the dialog counts and the elapsed time.
        root("dialogs_after") = preNDlg
        root("dialogs_leaked") = 0
        root("elapsedMs") = GetTickCount() - t0
        ToolMakeProject = JsonStringify(root)
        Exit Function
    End If
    ' ---- TASK-P1-125 (2.1): MOVE the old artifact aside, never delete it --------
    Dim fsoMP As Object
    Set fsoMP = CreateObject("Scripting.FileSystemObject")
    Dim sidecar As String
    sidecar = ""
    Dim stashed As Boolean
    stashed = False
    If Len(targetExe) > 0 Then
        If CBool(before("exists")) Then
            sidecar = targetExe & ".before-build"
            On Error Resume Next
            Err.Clear
            If fsoMP.FileExists(sidecar) Then fsoMP.DeleteFile sidecar, True
            Err.Clear
            fsoMP.MoveFile targetExe, sidecar
            If Err.Number = 0 Then stashed = True
            Err.Clear
            On Error GoTo 0
            If stashed Then root("artifact_backup_path") = sidecar
        End If
        ' A stale type library also makes VB6 re-prompt for a target. .exp/.lib are
        ' compiler INTERMEDIATES regenerated by every build, not deliverables, so they
        ' are still cleared; only the artifact itself is preserved.
        If LCase$(Right$(targetExe, 4)) = ".dll" Or LCase$(Right$(targetExe, 4)) = ".ocx" Then
            On Error Resume Next
            Dim stemMP As String
            stemMP = Left$(targetExe, Len(targetExe) - 4)
            If fsoMP.FileExists(stemMP & ".exp") Then fsoMP.DeleteFile stemMP & ".exp", True
            If fsoMP.FileExists(stemMP & ".lib") Then fsoMP.DeleteFile stemMP & ".lib", True
            On Error GoTo 0
        End If
    End If
    Dim compiled As Boolean
    compiled = False
    Dim errMsg As String
    errMsg = ""
    Dim attempt As Long
    Dim wdArmed As Boolean, wdArmedThis As Boolean, wdClicked As Boolean, wdStatus As String
    wdArmed = False
    wdArmedThis = False
    wdClicked = False
    wdStatus = ""
    For attempt = 1 To 3
        On Error Resume Next
        ' ALWAYS pin the target: an empty BuildFileName is exactly what makes
        ' VB6 open its Make Project file dialog instead of compiling.
        If Len(targetExe) > 0 Then p.BuildFileName = targetExe
        ' TASK-P1-125 (MCP-01 / 2.3): arm the one-shot confirm BEFORE the Make.
        ' MakeCompiledFile can put up a modal (the Make Project box, or a VB error
        ' box) and then BLOCK inside that modal's own message loop, so code that
        ' runs AFTER it returns can never click the button: the call hangs for ever
        ' and every later write tool answers reason=busy. Measured live while
        ' building this card: make_project stuck at 123 s with the IDE's own error
        ' box up (an English OK button), and read_active_dialogs got the busy
        ' receipt from inside that loop. The armed confirm's WM_TIMER IS dispatched
        ' by that modal loop (TASK-P1-084), which is the only place the click can
        ' come from. WoConfirmArm refuses to arm while a dialog is already up, so it
        ' can only ever press a button of a box THIS call opened.
        wdArmedThis = WinOps.WoConfirmArm(60000)
        If wdArmedThis Then wdArmed = True
        p.MakeCompiledFile
        If Err.Number = 0 Then
            compiled = True
        Else
            errMsg = Err.Description
        End If
        On Error GoTo 0
        wdClicked = wdClicked Or WinOps.WoConfirmClicked()
        wdStatus = WinOps.WoConfirmStatus()
        Call WinOps.WoConfirmDisarm
        If compiled Then Exit For
        If attempt < 3 Then Sleep 500
    Next
    root("compiled") = compiled
    ' ---- TASK-P1-125: is VB6's own Make box up? --------------------------------
    Dim risk As String, riskReason As String, riskClass As String, riskDetails As String
    Dim nDlg As Long
    nDlg = WinOps.HostModalRiskProbe("#32770", risk, riskReason, riskClass, riskDetails)
    root("host_modal_risk") = risk
    root("host_modal_risk_reason") = riskReason
    root("host_modal_risk_class") = riskClass
    root("modalDialogs") = riskDetails
    Dim buildStarted As Boolean
    Dim usedMenuFallback As Boolean
    buildStarted = False
    usedMenuFallback = False
    If Not compiled Then
        ' fallback: File > Make ... menu command (English or localized captions)
        Dim ctrl As Object
        Set ctrl = FindCmd("make", "File|Menu Bar|" & ChrW$(&H6587) & ChrW$(&H4EF6) & "|" & ChrW$(&H83DC) & ChrW$(&H5355) & ChrW$(&H6761), "Make|" & ChrW$(&H751F) & ChrW$(&H6210))
        If ctrl Is Nothing Then
            Err.Raise vbObjectError + 40, "CodeTools", "Could not compile: " & errMsg
        End If
        ctrl.Execute
        usedMenuFallback = True
        compiled = True
        Sleep 500
        nDlg = WinOps.HostModalRiskProbe("#32770", risk, riskReason, riskClass, riskDetails)
        root("host_modal_risk") = risk
        root("host_modal_risk_reason") = riskReason
        root("host_modal_risk_class") = riskClass
        root("modalDialogs") = riskDetails
    End If
    ' ---- TASK-P1-125 (2.3): confirm the box THIS call opened, then say so ------
    ' Confirming VB6's Make box inside make_project is legitimate: the caller asked
    ' for a compile and expects the call to block for it, and the box is reported as
    ' host_modal_risk=known so the blocking is explicit instead of silent. The defect
    ' was a tool the caller expected to return quickly (close_active_dialogs) doing it.
    If nDlg > 0 Then
        Sleep 400
        root("modalConfirmAttempted") = True
        ConfirmMakeDialog
        buildStarted = True
        If risk = "known" Then
            root("modal_note") = "VB6's own Make Project box was up and WAS confirmed: this call is now inside the host's modal loop for the whole compile. That is expected for make_project, and it is exactly why close_active_dialogs must never click this box."
        End If
    ElseIf compiled Then
        buildStarted = True
    ElseIf usedMenuFallback Then
        Sleep 400
        root("modalConfirmAttempted") = True
        ConfirmMakeDialog
        buildStarted = True
    End If
    root("buildStarted") = buildStarted
    root("usedMenuFallback") = usedMenuFallback
    ' ---- TASK-P1-125 (2.4): wait (bounded) for the artifact to land ------------
    Dim after As Object
    Set after = ArtifactFullSnap(targetExe)
    Dim waited As Long
    waited = 0
    Dim ceText As String
    Dim hasCompileErr As Boolean
    hasCompileErr = False
    Do While buildStarted And waited < MAKE_LAND_WAIT_MS
        hasCompileErr = WinOps.CompileErrorTextFromDialogs(ceText)
        If hasCompileErr Then Exit Do
        If CBool(after("exists")) Then
            If Not CBool(before("exists")) Then Exit Do
            If CStr(after("sha256")) <> CStr(before("sha256")) Then Exit Do
            If CDbl(after("mtime")) <> CDbl(before("mtime")) Then Exit Do
        End If
        Sleep 250
        waited = waited + 250
        Set after = ArtifactFullSnap(targetExe)
    Loop
    ' ---- TASK-P1-125 (2.1): if the host removed the artifact, put it back ------
    ' ---- TASK-P1-126 (F4): PUT THE ARTIFACT BACK, and PROVE it -------------------
    ' The P1-125 remedy only ran when THIS call had stashed a copy, and the verifier
    ' reproduced two independent constructions where the artifact ended up missing and
    ' was NOT restored. The restore below runs on every path, tries every candidate,
    ' verifies the restored bytes against the PRE-CALL digest, and fails LOUDLY with a
    ' reason when it cannot.
    Dim restored As Boolean
    restored = False
    Dim restoreAttempted As Boolean
    restoreAttempted = False
    Dim restoreSource As String
    restoreSource = ""
    Dim restoreFailed As String
    restoreFailed = ""
    Dim beforeShaA As String
    ' TASK-P1-127 (R1): retention + restore-basis evidence. Option Explicit is in force
    ' in this class, so every one of these needs its own declaration.
    Dim restoreBasis As String
    Dim restoreSourceSha As String
    Dim restoreMatchesPrecall As Boolean
    Dim retainedCandidatePath As String
    Dim retainedCandidateSha As String
    Dim retainedCandidateErr As String
    beforeShaA = CStr(before("sha256"))
    If Not CBool(after("exists")) Then
        Dim cands As Object
        Set cands = New Collection
        If Len(sidecar) > 0 Then cands.Add sidecar
        cands.Add targetExe & ".before-build"
        cands.Add targetExe & ".locked-build"
        Dim ci As Long
        For ci = 1 To cands.Count
            Dim cnd As String
            cnd = CStr(cands(ci))
            On Error Resume Next
            Err.Clear
            If fsoMP.FileExists(cnd) Then
                restoreAttempted = True
                ' TASK-P1-127 (R1): take the candidate's digest BEFORE it is moved, so
                ' the receipt can prove WHICH image was put back.
                restoreSourceSha = SnapSha(ArtifactFullSnap(cnd))
                If fsoMP.FileExists(targetExe) Then fsoMP.DeleteFile targetExe, True
                fsoMP.MoveFile cnd, targetExe
                If Err.Number = 0 Then
                    restored = True
                    restoreSource = cnd
                    Err.Clear
                    Exit For
                Else
                    restoreFailed = "restore failed: moving " & cnd & " back to " & targetExe & " raised error " & CStr(Err.Number) & " (" & Err.Description & ")"
                End If
            End If
            On Error GoTo 0
        Next ci
        If (Not restored) And Len(restoreFailed) = 0 Then
            restoreFailed = "restore failed: no candidate copy was found on disk (" & CStr(cands.Count) & " candidate path(s) checked: " & CStr(cands(1)) & " ...). Neither the pre-call image (moved aside by this call) nor the candidate retained by the PREVIOUS call (<exe>.before-build) exists on disk, so there was nothing to put back. The artifact was already missing when this call started."
        End If
        Set after = ArtifactFullSnap(targetExe)
        If restored Then
            Dim afterShaNow As String
            afterShaNow = SnapSha(after)
            ' With a pre-call image the restored bytes MUST equal it byte for byte.
            ' Without one the basis is the retained candidate, and it is reported as such.
            If Len(beforeShaA) > 0 Then
                restoreBasis = "pre_call_artifact"
                If Len(afterShaNow) = 0 Then
                    restored = False
                    restoreBasis = ""
                    restoreFailed = "restore NOT verified: the restored file's sha256 could not be read, so it cannot be proven equal to the pre-call digest (before=" & beforeShaA & ")"
                ElseIf afterShaNow <> beforeShaA Then
                    restored = False
                    restoreBasis = ""
                    restoreFailed = "restore NOT verified: the restored file does not match the pre-call digest (before=" & beforeShaA & " after=" & afterShaNow & ")"
                End If
            Else
                restoreBasis = "retained_candidate"
            End If
            If restored And Len(beforeShaA) > 0 Then restoreMatchesPrecall = (afterShaNow = beforeShaA)
        End If
    End If
    ' ---- TASK-P1-127 (R1): the candidate is RETAINED across calls -------------
    ' P1-126 only had a copy when THIS call happened to move the artifact aside
    ' first, so a deletion that landed before this call's own snapshot left nothing
    ' to put back (the verifier's Z5: the probe removed <exe> in a tight loop from the
    ' instant the call left the client, so the pre-call snapshot was already gone).
    ' The sidecar is now a RETAINED candidate: whenever a call ends with an artifact
    ' on disk, its bytes are COPIED (never moved) to <exe>.before-build, so the NEXT
    ' call can put the artifact back even when it is removed before that call can look
    ' at it. Nothing here ever overwrites a live artifact: the copy is read back only
    ' by the restore block above, and only when the artifact is missing.
    If CBool(after("exists")) And Not restored Then
        Dim retPath As String
        retPath = targetExe & ".before-build"
        On Error Resume Next
        Err.Clear
        If fsoMP.FileExists(retPath) Then fsoMP.DeleteFile retPath, True
        Err.Clear
        fsoMP.CopyFile targetExe, retPath
        If Err.Number = 0 Then
            retainedCandidatePath = retPath
            retainedCandidateSha = SnapSha(ArtifactFullSnap(retPath))
        Else
            retainedCandidateErr = CStr(Err.Number) & " (" & Err.Description & ")"
        End If
        Err.Clear
        On Error GoTo 0
    ElseIf restored Then
        ' the candidate was consumed by putting it back: nothing is retained now
        retainedCandidatePath = ""
        retainedCandidateSha = ""
    End If
    Set root("output") = after
    root("artifact_restored") = restored
    ' TASK-P1-126 (F4): machine-judgeable restore evidence, on every path.
    root("restore_attempted") = restoreAttempted
    root("restore_source") = restoreSource
    root("restore_failed_reason") = restoreFailed
    ' TASK-P1-127 (R1): VB6's IIf() evaluates BOTH branches, so the P1-126 form
    ' IIf(restored, CStr(after("sha256")), "") still touched the digest on the
    ' not-restored path (where the key is absent and reads back as Empty). Assign
    ' explicitly so every path publishes the key without evaluating a missing one.
    root("restore_verified_sha256") = ""
    If restored Then root("restore_verified_sha256") = SnapSha(after)
    ' TASK-P1-127 (R1): WHERE the restored bytes came from, and whether they are
    ' PROVEN equal to the pre-call image. restore_basis=pre_call_artifact means this
    ' call held the pre-call digest and verified against it; =retained_candidate means
    ' the artifact was already gone when the call started, so the bytes are the copy
    ' retained at the end of the previous call and no pre-call digest existed here.
    root("restore_basis") = restoreBasis
    root("restore_source_sha256") = restoreSourceSha
    root("restore_matches_precall_sha256") = restoreMatchesPrecall
    root("retained_candidate_path") = retainedCandidatePath
    root("retained_candidate_sha256") = retainedCandidateSha
    If Len(retainedCandidateErr) > 0 Then root("retained_candidate_error") = retainedCandidateErr
    root("restore_note") = "artifact_restored=true means the artifact was put back FROM a candidate copy and its sha256 matches the digest this call could verify (restore_basis names which: pre_call_artifact = the pre-call digest, retained_candidate = the copy kept at the end of the previous call). A false value always carries restore_failed_reason; it is never merely reported. <exe>.before-build is retained between calls so a later deletion can still be undone."
    ' ---- artifactChanged: the only honest "did anything get built" ------------
    Dim artChanged As Boolean
    artChanged = False
    If CBool(after("exists")) Then
        If Not CBool(before("exists")) Then
            artChanged = True
        ElseIf CDbl(after("size")) <> CDbl(before("size")) Then
            artChanged = True
        ElseIf CDbl(after("mtime")) <> CDbl(before("mtime")) Then
            artChanged = True
        ElseIf CStr(after("sha256")) <> CStr(before("sha256")) Then
            artChanged = True
        End If
    End If
    root("artifactChanged") = artChanged
    ' ---- TASK-P1-125 (3.2): is the artifact held by a running process? ---------
    Dim holderPid As Long
    holderPid = 0
    Dim lockerPid As Long
    lockerPid = 0
    If Len(targetExe) > 0 And (Not artChanged) Then lockerPid = ArtifactHolderPid(targetExe)
    Dim writeLocked As Boolean
    writeLocked = False
    If Len(targetExe) > 0 And (Not artChanged) Then writeLocked = ArtifactWriteLocked(targetExe)
    If lockerPid > 0 Then
        root("artifact_locked") = True
        root("artifact_locker_pid") = lockerPid
        root("artifact_holder_pid") = lockerPid
    ElseIf writeLocked Then
        root("artifact_locked") = True
    ElseIf CBool(before("exists")) And Not stashed Then
        ' the caller's previous artifact could not even be moved aside
        root("artifact_locked") = True
    End If
    If CBool(root("artifact_locked")) Then
        root("artifact_locked_advice") = "do NOT just retry: rename the existing artifact first (e.g. <name>.locked-build) or close the process holding it, then rebuild. artifact_locker_pid = 0 means the OS did not name the locker (Windows exposes no lock-owner API); it does NOT mean the file is unlocked."
        If lockerPid = 0 Then root("artifact_locker_pid_note") = "the OS did not name a lock owner; artifact_locked was decided by the exclusive-open test and/or by the artifact not being movable"
    End If
    ' ---- pending: pending=true must say HOW to confirm -------------------------
    Dim pending As Boolean
    pending = False
    If buildStarted And (Not CBool(after("exists"))) And (Not hasCompileErr) Then pending = True
    root("pending") = pending
    If pending Then
        root("pending_how_to_confirm") = "the artifact has not appeared yet: re-read output.exists / output.size / output.sha256 / output.mtime (this call already waited " & CStr(waited) & " ms) and check that the file starts with a PE header (MZ at offset 0). Do not read compiled/ok as evidence."
    End If
    ' ---- the in-IDE route has no compiler log: say so, do not fake one --------
    root("logAvailable") = False
    root("logUnavailableReason") = "the in-IDE Make route writes NO compiler log, so log/errorCount/firstError cannot come from a file here. Use route=cli (or the build tool) - an isolated VB6.EXE /make <vbp> /out <log> subprocess - to get log, errorCount and firstError{file,line,text} with a real file location."
    If Len(outPath) > 0 Then root("logPath") = outPath
    If hasCompileErr Then
        root("errorCount") = 1
        root("firstErrorSource") = "ide_compile_error_dialog"
        Dim fe2 As Object
        Set fe2 = CreateObject("Scripting.Dictionary")
        fe2("file") = ""
        fe2("line") = 0
        fe2("lineSpace") = "disk"
        fe2("text") = ceText
        Set root("firstError") = fe2
        root("note") = "VB6's own compile-error box was read while it was up (errorCount=1 because one box was seen, NOT because exactly one error exists); the in-IDE route has no log to count. route=cli gives a real count with file+line."
    ElseIf Not compiled Then
        root("note") = "MakeCompiledFile raised (last error: " & errMsg & ") and no compile-error box was readable; retry with route=cli to get a compiler log with firstError{file,line,text}"
    End If
    ' ---- verdict + ok: compiled alone is never a verdict ----------------------
    If CBool(root("modal_blocked")) Then
        root("verdict") = "modal_blocked"
    ElseIf pending Then
        root("verdict") = "pending"
    ElseIf artChanged Then
        root("verdict") = "built"
    ElseIf CBool(after("exists")) And (compiled Or buildStarted) Then
        root("verdict") = "compiled-no-artifact-change"
        root("artifactLockedHint") = "the Make returned without error but the artifact was not rewritten: it is most likely held open by another process (close it), or the build wrote identical bytes"
        root("artifactNote") = "treat this as NOT built: artifactChanged is false"
    ElseIf restored Then
        root("verdict") = "restored-previous-artifact"
    ElseIf Not compiled Then
        root("verdict") = "compile_failed"
    Else
        root("verdict") = "no-artifact-change"
    End If
    root("ok") = (artChanged And Not CBool(root("modal_blocked")))
    root("elapsedMs") = GetTickCount() - t0
    If Len(root("note")) = 0 Then
        root("note") = "one build reports only the first error; compiled:true is not evidence - judge artifactChanged / verdict / output{}"
    End If
    If restored Then root("artifact_backup_path") = ""
    If CBool(after("exists")) And stashed And Not restored Then root("artifact_backup_path") = ""
    ' ---- TASK-P1-125 (3): shape-aligned top-level location fields ----------------
    Dim feR As Object
    Set feR = root("firstError")
    If Not feR Is Nothing Then
        root("file") = CStr(feR("file"))
        root("reportedLine") = CLng(feR("line"))
        root("diskLine") = CLng(feR("line"))
    End If
    ' ---- TASK-P1-125 (3): do not LEAK dialogs -----------------------------------
    ' The A-side run measured a failed in-IDE Make leaving one MORE pair of boxes up
    ' per attempt (count grew to 7 over 5 failures), so the IDE stayed blocked. A
    ' call must not leave the IDE with more boxes than it found: the boxes THIS
    ' call created are dismissed before it returns, and the counts are machine
    ' readable. A Make box is cancelled, never confirmed - cancelling cannot start
    ' the compile this path decided not to run.
    Dim dRisk2 As String, dReason2 As String, dClass2 As String, dDetails2 As String
    Dim dlAfter As Long
    dlAfter = WinOps.HostModalRiskProbe("#32770", dRisk2, dReason2, dClass2, dDetails2)
    Dim dlLeaked As Long
    dlLeaked = dlAfter - preNDlg
    If dlLeaked < 0 Then dlLeaked = 0
    Dim dlAttempts As Long
    dlAttempts = 0
    ' TASK-P1-126 (F6): dismiss on EVERY path, pending included. The P1-125 guard
    ' skipped the dismissal while pending=true, and the verifier measured exactly one
    ' leftover box on the first attempt. A call must not leave the IDE with MORE boxes
    ' than it found: dialogs_after <= dialogs_before is the contract, and the compile
    ' error text (if any) is captured BEFORE any box is closed.
    If dlLeaked > 0 Then
        Dim dlIgnored As String
        dlIgnored = WinOps.CloseActiveDialogsJson(0, "Cancel", False, False)
        dlAttempts = 1
        dlIgnored = WinOps.CloseActiveDialogsJson(0, "", False, False)
        dlAfter = WinOps.HostModalRiskProbe("#32770", dRisk2, dReason2, dClass2, dDetails2)
        dlLeaked = dlAfter - preNDlg
        If dlLeaked < 0 Then dlLeaked = 0
    End If
    root("make_watchdog_armed") = wdArmed
    root("make_watchdog_clicked") = wdClicked
    root("make_watchdog_status") = wdStatus
    If wdArmed Then
        root("make_watchdog_note") = "the one-shot confirm was armed across MakeCompiledFile so a modal VB6 opened INSIDE that call could still be dismissed (its WM_TIMER is dispatched by the modal loop). armed=" & CStr(wdArmed) & " clicked=" & CStr(wdClicked) & " status=" & wdStatus
    End If
    Call WinOps.WoConfirmDisarm
    root("dialogs_after") = dlAfter
    root("dialogs_leaked") = dlLeaked
    root("dialogs_close_attempts") = dlAttempts
    ' TASK-P1-126 (F6): the net-zero contract, as a boolean the judge can read.
    root("dialogs_before") = preNDlg
    root("dialogs_net_zero") = (dlAfter <= preNDlg)
    If Not CBool(root("dialogs_net_zero")) Then
        root("dialogs_note_net_zero") = "this call still left " & CStr(dlLeaked) & " more box(es) up than it found after " & CStr(dlAttempts) & " dismissal pass(es); the captions may be outside the safe whitelist - read them with read_active_dialogs and click explicitly"
    End If
    If dlAttempts > 0 Then
        root("dialogs_note") = "this call left " & CStr(dlLeaked) & " more box(es) up than it found and dismissed them itself (Cancel first, then safe captions); any compile-error text was captured into firstError.text before the box was closed, so nothing is lost by dismissing it"
        If Len(root("dialogs_note")) > 0 And Len(CStr(feR("text"))) > 0 Then root("firstErrorSource") = "ide_compile_error_dialog"
    End If
    ' ---- TASK-P1-126 (F5): the in-IDE route STRUCTURALLY cannot give the root cause
    ' The verifier measured ide-route receipts with errorCount=0, firstError empty and
    ' logAvailable=false, and re-tested with an explicit route=cli (exit 0, no log).
    ' VBProject.MakeCompiledFile reports only a boolean, so there is NO compiler
    ' subprocess and therefore no log to read on this route - that is structural, not a
    ' transient gap. Say so, and name the executable alternative.
    If Not CBool(root("logAvailable")) Then
        root("diagnosisAvailable") = False
        root("diagnosisUnavailableReason") = "the in-IDE Make route has NO compiler subprocess, so no /out log and no compiler error text can exist for it - structurally, not as a failure of this call"
        root("diagnosisAlternative") = "call build (or make_project with route=cli): it runs an isolated VB6.EXE /make <vbp> /out <log> and returns log plus firstError{file,line,text}. Measured caveat: while ANOTHER VB6 instance holds the same .vbp, /make exits 0 without writing the log - close that instance, or build a copy of the project, and retry."
        root("suggestedAction") = CStr(root("diagnosisAlternative"))
    Else
        root("diagnosisAvailable") = True
        root("diagnosisUnavailableReason") = ""
        root("diagnosisAlternative") = ""
    End If
    ' ---- TASK-P1-132 (A10): LAA on the artifact (sidecar aware) ----------------
    If laaUsable Then
        ApplyLaaInto root, targetExe, laaWant
    Else
        LaaConfigRefusal root, laaProjFile, laaCfgErr
    End If
    root("laaSource") = laaSrc
    root("laaProjectFile") = laaProjFile
    root("laaSourceNote") = laaSrcNote
    root("laaConfigError") = laaCfgErr
    ' ---- TASK-P1-132 (MCP-lead-2 item 1): compiled-but-nothing-produced must
    ' never read like a success ------------------------------------------------
    If (Not artChanged) And compiled Then
        root("unbuiltReasonUnknown") = True
        root("artifactNote") = "NOT built: the artifact was neither produced nor rewritten, and the in-IDE Make route cannot say WHY - a compile failure is possible (compiled=true only means MakeCompiledFile returned without raising). Read diagnosis{} below."
    End If
    ' ---- TASK-P1-132 (MCP-lead-2 item 2): isolated-copy diagnosis -------------
    Dim diagReason As String
    diagReason = ""
    If diagnoseWant Then
        diagReason = "requested"
    ElseIf (Not artChanged) And (compiled Or buildStarted) Then
        diagReason = "compiled-but-no-artifact-change"
    End If
    If Len(diagReason) > 0 Then
        DiagnoseInto root, p.FileName, timeoutMs, diagReason
        Dim dg As Object
        Set dg = root("diagnosis")
        If CBool(dg("attempted")) Then
            root("diagnosisRoute") = CStr(dg("route"))
            root("diagnosisTmpDir") = CStr(dg("tmpDir"))
            root("diagnosisCopiedFiles") = CLng(dg("copiedFiles"))
            root("diagnosisMissingFiles") = CStr(dg("missingFiles"))
            root("diagnosisError") = CStr(dg("error"))
            root("diagnosisAvailable") = True
            root("diagnosisUnavailableReason") = ""
            root("diagnosisAlternative") = ""
            If CBool(dg("ok")) Then
                root("unbuiltReasonUnknown") = False
                root("artifactNote") = "the artifact was not rewritten, but an ISOLATED compile of a COPY of this project SUCCEEDED with 0 errors - the source is valid, so the in-IDE Make simply did not rewrite the artifact. See diagnosis{}."
            End If
            If Len(CStr(dg("log"))) > 0 Then
                root("logAvailable") = True
                root("logWrittenByThisCall") = True
                root("log") = CStr(dg("log"))
                root("logPath") = CStr(dg("logPath"))
                root("logBytes") = CLng(dg("logBytes"))
                root("errorCount") = CLng(dg("errorCount"))
                Dim dfef As Object
                Set dfef = dg("firstError")
                If Not dfef Is Nothing Then
                    Set root("firstError") = dfef
                    root("firstErrorSource") = "isolated_copy_make"
                    root("file") = CStr(dfef("file"))
                    root("reportedLine") = CLng(dfef("line"))
                    root("diskLine") = CLng(dfef("line"))
                End If
                root("diagnosisNote") = "log/firstError come from the ISOLATED copy compile, not from the in-IDE Make (which has no compiler subprocess). The copy is the state ON DISK: unsaved pane edits are not in it."
            Else
                root("diagnosisNote") = "the isolated /make wrote no readable log; see diagnosis{} and diagnosisError"
            End If
        End If
    End If
    ToolMakeProject = JsonStringify(root)
End Function

Private Sub ConfirmMakeDialog()
    ' Click the default button of any #32770 dialog owned by THIS process.
    ' Matching by window class + process id covers every caption and UI
    ' language (Make Project / overwrite prompt / VB6 error box); the old
    ' exact-title match on the Chinese caption did not.
    Const BM_CLICK As Long = &HF5
    Const GW_CHILD As Long = 5
    Const GW_HWNDNEXT As Long = 2
    Const GWL_STYLE As Long = -16
    Const BS_DEFPUSHBUTTON As Long = &H1
    On Error Resume Next
    Dim myPid As Long
    myPid = GetCurrentProcessId()
    Dim hDlg As Long, hChild As Long, pid As Long, btn As Long, style As Long
    hDlg = FindWindowExW(0, 0, StrPtr("#32770"), 0)
    Do While hDlg <> 0
        pid = 0
        GetWindowThreadProcessId hDlg, pid
        If pid = myPid Then
            btn = 0
            hChild = GetWindow(hDlg, GW_CHILD)
            Do While hChild <> 0
                style = GetWindowLongW(hChild, GWL_STYLE)
                If (style And BS_DEFPUSHBUTTON) <> 0 Then
                    btn = hChild
                    Exit Do
                End If
                hChild = GetWindow(hChild, GW_HWNDNEXT)
            Loop
            If btn <> 0 Then PostMessageW btn, BM_CLICK, 0, 0
        End If
        hDlg = FindWindowExW(0, hDlg, StrPtr("#32770"), 0)
    Loop
End Sub

Private Function ToolControlExecution(ByVal mode As String) As String
    Dim m As String
    m = LCase$(mode)
    ' Card TASK-P1-005 (P4 D7): inputSchema advertises "run (start, or continue when
    ' already broken)", but the code took the design path, whose mandatory pre-run
    ' "save the project" step the IDE REFUSES in break state (measured:
    '   拒绝的权限), so run was unusable exactly where it was documented to work.
    ' While broken the menu item is Continue, so go straight there.
    If m = "run" Then
        If GetIdeState() = "break" Then m = "continue"
    End If
    If m = "run" Or m = "full_compile" Or m = "fullcompile" Or m = "restart" Then
        ' Running an unsaved project makes VB6 pop a modal Save dialog which
        ' blocks the IDE and the MCP server. Saving only the .vbp is NOT enough:
        ' a component that has never been written to disk still makes the IDE
        ' open "文件另存为" (measured 2026-09-21: the run stalls with a modal
        ' dialog and the caller sees only "mode=run, state=design"), so flush
        ' everything through the same path save_project uses.
        Dim sp As VBIDE.VBProject
        Set sp = GetProject()
        If Len(sp.FileName) = 0 Then
            Err.Raise vbObjectError + 41, "CodeTools", "Project has no file name: running would open VB6's Save dialog. Call save_project with a path first, then retry."
        End If
        Dim saveJson As String
        Dim svErr As Long
        Dim svDesc As String
        saveJson = ""
        svErr = 0
        svDesc = ""
        On Error Resume Next
        saveJson = ToolSaveProject("")
        svErr = Err.Number
        svDesc = Err.Description
        Err.Clear
        On Error GoTo 0
        If svErr <> 0 Then Err.Raise vbObjectError + 41, "CodeTools", "Could not save the project before running: " & svDesc
        OpLog "control_execution " & m & ": pre-run flush " & Left$(saveJson, 200)
    End If
    Dim ctrl As Object
    Dim execCount As Long
    Dim whyExec As String
    execCount = 0
    ' TASK-P1-029: remember the state BEFORE the command so the response can prove a
    ' transition instead of reporting a stale snapshot.
    Dim beforeState As String
    beforeState = GetIdeState()
    If m = "restart" Then
        Set ctrl = FindCmd("end", "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C), "End|" & ChrW$(&H7ED3) & ChrW$(&H675F))
        If ctrl Is Nothing Then Set ctrl = m_cachedEnd Else Set m_cachedEnd = ctrl
        If ctrl Is Nothing Then Err.Raise vbObjectError + 41, "CodeTools", "Menu command not found: End"
        execCount = execCount + ExecIfEnabled(ctrl, m)
        ' the End teardown completes on the IDE message pump only after
        ' this response is sent; arm a deferred Start instead of
        ' clicking Start right away (it would still be disabled)
        If Not m_server Is Nothing Then m_server.RequestRestart
    Else
        Dim cands As String
        Dim barCands As String
        Select Case m
            Case "run"
                ' Start (design state) or Continue (already running or broken):
                ' the localized IDE renames the item, so try both captions.
                cands = "Start|" & ChrW$(&H542F) & ChrW$(&H52A8) & "|Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "continue", "resume"
                cands = "Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "step_into", "stepinto"
                cands = "Step Into|" & ChrW$(&H9010) & ChrW$(&H8BED) & ChrW$(&H53E5)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "step_over", "stepover"
                cands = "Step Over|" & ChrW$(&H9010) & ChrW$(&H8FC7) & ChrW$(&H7A0B)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "step_out", "stepout"
                cands = "Step Out|" & ChrW$(&H8DF3) & ChrW$(&H51FA)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "run_to_cursor", "runtocursor"
                cands = "Run To Cursor|" & ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H5230) & ChrW$(&H5149) & ChrW$(&H6807) & ChrW$(&H5904)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "full_compile", "fullcompile"
                cands = "Start With Full Compile|" & ChrW$(&H5168) & ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H6267) & ChrW$(&H884C)
                barCands = "Run|Menu Bar|" & ChrW$(&H8FD0) & ChrW$(&H884C) & "|" & ChrW$(&H83DC) & ChrW$(&H5355) & ChrW$(&H6761)
            Case "break"
                cands = "Break|" & ChrW$(&H4E2D) & ChrW$(&H65AD)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case "end", "stop"
                cands = "End|" & ChrW$(&H7ED3) & ChrW$(&H675F)
                barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
            Case Else
                Err.Raise vbObjectError + 41, "CodeTools", "Invalid mode: " & mode
        End Select
        Set ctrl = FindCmd(m, barCands, cands)
        If ctrl Is Nothing Then
            Select Case m
                Case "run": Set ctrl = m_cachedStart
                Case "break": Set ctrl = m_cachedBreak
                Case "end", "stop": Set ctrl = m_cachedEnd
                Case "full_compile", "fullcompile": Set ctrl = m_cachedFull
                Case "step_into", "stepinto": Set ctrl = m_cachedStepInto
                Case "step_over", "stepover": Set ctrl = m_cachedStepOver
                Case "step_out", "stepout": Set ctrl = m_cachedStepOut
                Case "run_to_cursor", "runtocursor": Set ctrl = m_cachedRtc
            End Select
        End If
        ' Card TASK-P1-005 fallback 2: FindBarControl filters bars by the localized,
        ' state-dependent bar NAME and descends ONE popup level only. Walk every
        ' CommandBar recursively by caption before giving up.
        If ctrl Is Nothing Then Set ctrl = FindBarControlDeep(cands)
        If Not ctrl Is Nothing Then
            Select Case m
                Case "run": Set m_cachedStart = ctrl
                Case "break": Set m_cachedBreak = ctrl
                Case "end", "stop": Set m_cachedEnd = ctrl
                Case "full_compile", "fullcompile": Set m_cachedFull = ctrl
                Case "step_into", "stepinto": Set m_cachedStepInto = ctrl
                Case "step_over", "stepover": Set m_cachedStepOver = ctrl
                Case "step_out", "stepout": Set m_cachedStepOut = ctrl
                Case "run_to_cursor", "runtocursor": Set m_cachedRtc = ctrl
            End Select
        End If
        ' Card TASK-P1-005: 'continue' does not exist in the Debug menu in design
        ' state at all (the IDE shows Start, not Continue), so the lookup failed and
        ' the tool raised a bare "Menu command not found" - exactly the opaque symptom
        ' this card is about. Every mode now answers with the same structured failure
        ' instead of aborting, and names the captions that were searched.
        If ctrl Is Nothing Then
            Dim nf As Object
            Set nf = CreateObject("Scripting.Dictionary")
            nf("mode") = m
            nf("executed") = False
            nf("reason") = "not_found"
            nf("candidates") = cands
            nf("state") = GetIdeState()
            ToolControlExecution = JsonStringify(nf)
            Exit Function
        End If
        execCount = ExecIfEnabledWhy(ctrl, m, whyExec)
    End If
    ' break/end/stop: yield so the debug toolbar reflects the new
    ' state. run/full_compile/restart: DoEvents would block inside
    ' the running app's modal message loop and the RPC response
    ' would never be sent, so return immediately after Execute.
    ' single steps are short: yield so the IDE reflects the new line.
    ' run/continue/run_to_cursor may run the app for a long time - DoEvents
    ' would block inside the app message loop and the RPC response would
    ' never be sent, so those return immediately after Execute.
    ' TASK-P1-029: the Run/Debug menu command is processed ASYNCHRONOUSLY on the IDE
    ' message pump, so ONE DoEvents is not enough. Measured on a DoEvents loop: end
    ' answered {executed:true, state:"run"} while an independent ide_state probe ~2s
    ' later was already "design" - the command had worked, the receipt was stale. Poll
    ' for the transition (bounded) and never claim success when the state never moves.
    Dim expectChange As Boolean
    expectChange = (m = "break" Or m = "end" Or m = "stop" Or m = "restart")
    Dim afterState As String
    afterState = GetIdeState()
    If expectChange And execCount > 0 Then afterState = WaitIdeLeave(beforeState, 12000)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("mode") = m
    root("executed") = (execCount > 0)
    root("stateBefore") = beforeState
    ' Card TASK-P1-005: break answered {"executed":false} with no reason at all, so a
    ' caller could not tell "already broken" from "the item is disabled".
    If execCount = 0 Then
        If Len(whyExec) > 0 Then
            root("reason") = whyExec
        ElseIf m = "break" And GetIdeState() = "break" Then
            root("reason") = "already_broken"
        Else
            root("reason") = "not_executed"
        End If
    End If
    On Error Resume Next
    root("caption") = ctrl.Caption
    On Error GoTo 0
    If m = "restart" Then root("restartPending") = True
    root("state") = afterState
    ' TASK-P1-029 (receipt honesty, same discipline as the A batch): a menu item that was
    ' found and executed is NOT proof that the IDE changed state. If a command that must
    ' change the state did not, say so structurally instead of reporting executed:true.
    If expectChange And execCount > 0 And afterState = beforeState Then
        root("ok") = False
        root("reason") = "state_not_changed"
        root("before") = beforeState
        root("after") = afterState
        root("advice") = "the IDE did not leave state '" & beforeState & "' within 12s. With a tight loop that does not pump messages the command cannot be processed at all (measured: end/break/restart time out and the MCP endpoint becomes unreachable) - the only recovery is to stop the VB6 process. With a DoEvents loop the command normally works; re-probe with ide_state."
    End If
    ToolControlExecution = JsonStringify(root)
End Function

' TASK-P1-029: pump the message loop until the IDE leaves fromState (bounded).
Private Function WaitIdeLeave(ByVal fromState As String, ByVal timeoutMs As Long) As String
    Dim t0 As Long, s As String
    t0 = GetTickCount()
    s = GetIdeState()
    Do While s = fromState And (GetTickCount() - t0) < timeoutMs
        DoEvents
        ' TASK-P1-126 (MCP-08): the cooperative abort is honoured HERE, at a pump point
        ' of the in-flight call (a re-entrant ide_state{cancel_inflight:true} set the flag).
        MaybeHonorCancel
        s = GetIdeState()
    Loop
    WaitIdeLeave = s
End Function

' ---------------------------------------------------------------------
' native IDE breakpoint (Debug menu; does not modify any source file)
' ---------------------------------------------------------------------
Private Function BreakpointLocationProblem(ByVal comp As String, ByVal line As Long) As String
    ' Returns "" when the position can hold a breakpoint, else a short reason code.
    ' VB6 refuses breakpoints in "some regions": blank lines, comment lines and the
    ' declaration section. Asking anyway aimed the IDE at a position it will not
    ' accept, and the request died with the connection closed and no reply.
    On Error Resume Next
    If Len(comp) = 0 Then Exit Function
    Dim c As VBIDE.VBComponent
    Set c = FindComponent(comp)
    If c Is Nothing Then
        BreakpointLocationProblem = "component_not_found"
        Exit Function
    End If
    Dim cm As VBIDE.CodeModule
    Set cm = c.CodeModule
    If cm Is Nothing Then
        BreakpointLocationProblem = "no_code_module"
        Exit Function
    End If
    If line < 1 Or line > cm.CountOfLines Then
        BreakpointLocationProblem = "line_out_of_range"
        Exit Function
    End If
    If line <= cm.CountOfDeclarationLines Then
        BreakpointLocationProblem = "declaration_section"
        Exit Function
    End If
    Dim txt As String
    txt = Trim$(cm.Lines(line, 1))
    If Len(txt) = 0 Then
        BreakpointLocationProblem = "blank_line"
        Exit Function
    End If
    If Left$(txt, 1) = "'" Then
        BreakpointLocationProblem = "comment_line"
        Exit Function
    End If
    If LCase$(Left$(txt, 4)) = "rem " Or LCase$(txt) = "rem" Then
        BreakpointLocationProblem = "comment_line"
    End If
End Function

Private Function ToolNativeBreakpoint(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim barCands As String
    Dim ctrlCands As String
    barCands = "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C)
    Dim comp As String
    Dim line As Long
    comp = JsonGetStr(args, "component")
    line = JsonGetLong(args, "line", 0)
    ' Card: a breakpoint request aimed at a position VB6 will not accept used to
    ' take down the request (connection closed, no reply). Validate first and
    ' answer with a structured failure instead.
    Dim preReason As String
    preReason = ""
    Dim cp As VBIDE.CodePane
    ' TASK-P1-164 (R18-2 / R18-3): BEFORE anything moves, remember the caret and the pane
    ' text so the receipt can PROVE the pane was not rewritten and the caret was put back.
    Dim curOk As Boolean
    Dim curSL As Long, curSC As Long, curEL As Long, curEC As Long
    Dim pAssertComp As String, pAssertBefore As String
    Dim pAssertOn As Boolean
    Dim curRestored As Boolean
    Dim curNote As String
    Dim fgBefore As Long, fgAfter As Long
    curOk = False
    pAssertOn = False
    pAssertComp = ""
    pAssertBefore = ""
    curRestored = False
    curNote = "not attempted"
    fgBefore = WoFrontHwnd()
    Select Case action
        Case "toggle"
            ctrlCands = "Toggle Breakpoint|" & ChrW$(&H5207) & ChrW$(&H6362) & ChrW$(&H65AD) & ChrW$(&H70B9)
            If line > 0 Then
                preReason = BreakpointLocationProblem(comp, line)
                If Len(preReason) = 0 Then
                    Set cp = GetCodePane(comp)
                    On Error Resume Next
                    Err.Clear
                    If cp Is Nothing Then
                        preReason = "no_code_pane"
                    Else
                        ' R18-3: the IDE has NO breakpoint API, so "Toggle Breakpoint" acts on the
                        ' CARET (Debug menu item). Save the caret and the pane text here; both are
                        ' restored/checked after the menu command has run.
                        curSL = 0: curSC = 0: curEL = 0: curEC = 0
                        cp.GetSelection curSL, curSC, curEL, curEC
                        If Err.Number = 0 Then curOk = True
                        Err.Clear
                        pAssertComp = comp
                        pAssertBefore = PaneTextOfName(comp, pAssertOn)
                        cp.SetSelection line, 1, line, 1
                        If Err.Number <> 0 Then
                            preReason = "selection_failed"
                            Err.Clear
                        End If
                    End If
                    On Error GoTo 0
                End If
            End If
        Case "clear_all"
            ctrlCands = "Clear All Breakpoints|" & ChrW$(&H6E05) & ChrW$(&H9664) & ChrW$(&H6240) & ChrW$(&H6709) & ChrW$(&H65AD) & ChrW$(&H70B9)
            ' R18-2: clear_all names no component, so the assertion target is the ACTIVE code
            ' pane. When there is none, the assertion is reported as notApplicable rather than
            ' skipped silently. Clear All Breakpoints does not touch the code text either.
            pAssertComp = PaneCompOfActive()
            If Len(pAssertComp) > 0 Then pAssertBefore = PaneTextOfName(pAssertComp, pAssertOn)
        Case Else
            Err.Raise vbObjectError + 84, "CodeTools", "Invalid action: " & action & " (expected toggle/clear_all)"
    End Select
    If Len(preReason) > 0 Then
        Dim rej As Object
        Set rej = CreateObject("Scripting.Dictionary")
        rej("action") = action
        rej("executed") = False
        rej("reason") = preReason
        If Len(comp) > 0 Then rej("component") = comp
        If line > 0 Then rej("line") = line
        If Len(comp) > 0 And line > 0 Then
            ' ruling 4
            Set rej("coords") = CoordsDictOfName(comp, line, "ide")
        End If
        rej("state") = GetIdeState()
        rej("note") = "VB6 does not allow a breakpoint at this position (blank line, comment, declaration section, or out of range)."
        ToolNativeBreakpoint = JsonStringify(rej)
        Exit Function
    End If
    Dim ctrl As Object
    Set ctrl = FindCmd(action, barCands, ctrlCands)
    If ctrl Is Nothing Then
        Err.Raise vbObjectError + 84, "CodeTools", "Debug menu item not found for action=" & action & ". The CommandBars model is suspended while the program runs - stop it (end/break) and retry."
    End If
    Dim n As Long
    n = ExecIfEnabled(ctrl, action)
    DoEvents
    ' ---- R18-3 compensation: restore the caret, then give the foreground back --------
    ' VB6 exposes NO breakpoint API (see list_breakpoints), and the Debug menu item acts on
    ' the CARET, so there is NO channel that sets an IDE-native breakpoint "without moving
    ' the cursor and without changing text". What can be done - and is done here - is to
    ' put the caret back exactly where it was and to hand the foreground back to whatever
    ' window owned it. The receipt says all three facts so the trade-off is visible.
    If LCase$(action) = "toggle" Then
        If curOk Then
            On Error Resume Next
            Set cp = GetCodePane(comp)
            If cp Is Nothing Then
                curNote = "the code pane could not be resolved again, so the caret was NOT restored"
            Else
                Err.Clear
                cp.SetSelection curSL, curSC, curEL, curEC
                If Err.Number = 0 Then
                    curRestored = True
                    curNote = "the caret was put back at the selection read before the toggle"
                Else
                    curNote = "SetSelection raised " & CStr(Err.Number) & " (" & Err.Description & "); the caret was NOT restored"
                    Err.Clear
                End If
            End If
            On Error GoTo 0
        Else
            curNote = "the caret could not be read before the toggle, so nothing could be restored"
        End If
    Else
        curNote = "action " & action & " does not move the caret"
    End If
    fgAfter = WoFrontHwnd()
    If fgAfter <> fgBefore Then
        WoFrontRestore fgBefore
        fgAfter = WoFrontHwnd()
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("action") = action
    root("executed") = (n > 0)
    root("caption") = ctrl.Caption
    root("cursorRestored") = curRestored
    root("cursorNote") = curNote
    root("foregroundBefore") = fgBefore
    root("foregroundAfter") = fgAfter
    root("foregroundTookFocus") = (fgAfter <> fgBefore)
    root("foregroundReadable") = (fgBefore <> 0 Or fgAfter <> 0)
    ' TASK-P1-164 (A): a menu-driven toggle CANNOT be issued without activating the IDE, so
    ' quiet mode can only GIVE THE FOREGROUND BACK (done above). Report that honestly rather
    ' than claiming a suppression that did not happen.
    If QuietActive() Then
        root("quietModeSuppressedFocus") = False
        root("quietModeForegroundRestored") = (fgAfter = fgBefore)
        root("quietModeNote") = "quiet mode is ON, but a Debug-menu breakpoint toggle cannot be issued without activating the IDE: the foreground was taken and then GIVEN BACK (foregroundTookFocus / quietModeForegroundRestored say which). Quiet mode does not suppress this one - it only compensates for it."
    End If
    root("ideNativeBreakpointApi") = False
    root("breakpointWithoutCursorMove") = False
    root("breakpointQuietNote") = "VB6's extensibility model exposes NO breakpoint API (list_breakpoints says the same), and the Debug menu item acts on the caret; there is therefore NO channel that sets an IDE-native breakpoint without moving the cursor. This tool moves the caret to the target line and, as compensation, restores the previous selection afterwards (cursorRestored) and gives the foreground back (foregroundTookFocus). The alternative, set_breakpoint, does not move the caret but WRITES text (an injected Debug.Assert marker block)."
    If Len(comp) > 0 Then root("component") = comp
    If line > 0 Then root("line") = line
    If Len(comp) > 0 And line > 0 Then
        ' ruling 4
        Set root("coords") = CoordsDictOfName(comp, line, "ide")
    End If
    root("state") = GetIdeState()
    root("note") = "Menu-driven IDE breakpoint: no source file was modified. The IDE keeps breakpoints only for the current session."
    ' R18-2: the write-after assertion. A native breakpoint is menu-driven and writes no
    ' text, so the contract here is "the pane text is UNCHANGED"; a change flips the
    ' receipt to ok:false / reason:paneTextMismatch with the differing lines.
    PaneAssertFinish root, pAssertComp, pAssertBefore, pAssertOn, "native_breakpoint"
    ToolNativeBreakpoint = JsonStringify(root)
End Function

' ---- data access for the MCP resource layer (not tools) ----

Public Function ComponentNamesList() As Object
    Dim items As Object
    Set items = New Collection
    Dim p As VBIDE.VBProject
    Dim c As VBIDE.VBComponent
    On Error Resume Next
    Set p = GetProject()
    If p Is Nothing Then
        Set ComponentNamesList = items
        On Error GoTo 0
        Exit Function
    End If
    For Each c In p.VBComponents
        items.Add c.Name
    Next
    On Error GoTo 0
    Set ComponentNamesList = items
End Function

Public Function ComponentCodeText(ByVal name As String) As String
    On Error Resume Next
    Dim c As VBIDE.VBComponent
    Set c = FindComponent(name)
    If c Is Nothing Then Exit Function
    ' TASK-P1-062: guard the empty-component case (CountOfLines = 0 => Lines(1,0) = error 5)
    If c.CodeModule.CountOfLines > 0 Then ComponentCodeText = c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
End Function

Public Function ProjectVbpText() As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If p Is Nothing Then Exit Function
    ProjectVbpText = ProjShip.ReadTextFile(p.FileName)
End Function

' ---------------------------------------------------------------------
' offline VB tooling (VbFmt.bas): formatter and cross-project rename
' ---------------------------------------------------------------------
Private Function ToolFormatCode(ByVal args As Object) As String
    Dim path As String
    path = JsonGetStr(args, "path")
    If Len(path) = 0 Then Err.Raise vbObjectError + 81, "CodeTools", "path is required"
    Dim indentUnit As String
    indentUnit = JsonGetStr(args, "indent")
    If Len(indentUnit) = 0 Then indentUnit = "    "
    ToolFormatCode = VbFmt.FormatJson(ResolvePath(path), indentUnit, JsonGetBool(args, "apply", False))
End Function

Private Function ToolRenameSymbol(ByVal args As Object) As String
    Dim vbp As String
    vbp = JsonGetStr(args, "path")
    If Len(vbp) = 0 Then Err.Raise vbObjectError + 83, "CodeTools", "path is required (the .vbp)"
    Dim vbpFull As String
    vbpFull = ResolvePath(vbp)
    ' TASK-P1-110 (D-1107-05): a NONEXISTENT .vbp used to be reported as a successful 0-change dry
    ' run - ProjShip.ReadTextFile returns "" for a missing file, so the component list came back
    ' empty and nothing was found to rename. Refuse it, and say which paths were tried.
    If Not CtlFileExists(vbpFull) Then
        Dim nf As Object
        Set nf = CreateObject("Scripting.Dictionary")
        nf("ok") = False
        nf("reason") = "project_file_not_found"
        nf("tool") = "rename_symbol"
        nf("argument") = "path"
        nf("path") = vbp
        nf("resolvedPath") = vbpFull
        Dim nfTried As Object
        Set nfTried = New Collection
        nfTried.Add vbp
        If StrComp(vbp, vbpFull, vbTextCompare) <> 0 Then nfTried.Add vbpFull
        Set nf("triedPaths") = nfTried
        nf("message") = "project file not found: " & vbpFull
        nf("suggestedAction") = "pass the path of an EXISTING .vbp; a relative path resolves against the project directory, not the process working directory"
        nf("sideEffects") = "none: the call was refused before any file was read or written"
        ToolRenameSymbol = JsonStringify(nf)
        Exit Function
    End If
    ' TASK-P1-111 (S3): an EXISTING file that is not a VB6 project used to answer as a SUCCESSFUL
    ' 0-change dry run - ProjShip.ComponentPaths simply finds no Module=/Class= lines in it. This is
    ' the sibling of TASK-P1-110's project_file_not_found, but a DIFFERENT condition: the path does
    ' resolve to a real file, it just is not a .vbp (the declared contract of `path`). Refuse before
    ' anything is read or written.
    If LCase$(Right$(vbpFull, 4)) <> ".vbp" Then
        Dim naj As Object
        Set naj = CreateObject("Scripting.Dictionary")
        naj("ok") = False
        naj("reason") = "not_a_project_file"
        naj("tool") = "rename_symbol"
        naj("argument") = "path"
        naj("path") = vbp
        naj("resolvedPath") = vbpFull
        Dim najDot As Long
        Dim najExt As String
        najExt = ""
        najDot = InStrRev(vbpFull, ".")
        If najDot > 0 Then najExt = Mid$(vbpFull, najDot)
        naj("looksLike") = najExt
        Dim najTried As Object
        Set najTried = New Collection
        najTried.Add vbp
        If StrComp(vbp, vbpFull, vbTextCompare) <> 0 Then najTried.Add vbpFull
        Set naj("triedPaths") = najTried
        naj("message") = "not a VB6 project file (.vbp): " & vbpFull
        naj("suggestedAction") = "pass the path of the .vbp whose component files should be searched; to rename a COMPONENT (module/class/form) use rename_component"
        naj("sideEffects") = "none: the call was refused before any file was read or written"
        ToolRenameSymbol = JsonStringify(naj)
        Exit Function
    End If
    ToolRenameSymbol = VbFmt.RenameJson(vbpFull, JsonGetStr(args, "old_name"), JsonGetStr(args, "new_name"), _
                                        JsonGetBool(args, "dry_run", True), JsonGetBool(args, "skip_attributes", True))
End Function

' ---------------------------------------------------------------------
' project group / project snapshot / packaging / VCS (ProjShip.bas)
' ---------------------------------------------------------------------
Private Function Qq(ByVal s As String) As String
    Qq = Chr$(34) & s & Chr$(34)
End Function

Private Function ToolVbgGroup(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim raw As String
    raw = JsonGetStr(args, "path")
    If Len(raw) = 0 Then Err.Raise vbObjectError + 77, "CodeTools", "path is required (the .vbg project group)"
    Dim vbg As String
    vbg = ResolvePath(raw)
    Dim proj As String
    proj = JsonGetStr(args, "project")
    If Len(proj) > 0 Then proj = ResolvePath(proj)
    Select Case action
        Case "list"
            ToolVbgGroup = ProjShip.VbgListJson(vbg)
        Case "add"
            If Len(proj) = 0 Then Err.Raise vbObjectError + 77, "CodeTools", "project is required for action=add"
            ToolVbgGroup = ProjShip.VbgAddJson(vbg, proj)
        Case "remove"
            If Len(proj) = 0 Then Err.Raise vbObjectError + 77, "CodeTools", "project is required for action=remove"
            ToolVbgGroup = ProjShip.VbgRemoveJson(vbg, proj)
        Case Else
            Err.Raise vbObjectError + 77, "CodeTools", "Invalid action: " & action & " (expected list/add/remove)"
    End Select
End Function

Private Function ToolProjectSnapshot(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim rawPath As String
    rawPath = JsonGetStr(args, "path")
    Dim rawDir As String
    rawDir = JsonGetStr(args, "dir")
    Select Case action
        Case "create"
            If Len(rawPath) = 0 Then Err.Raise vbObjectError + 78, "CodeTools", "path is required for action=create"
            Dim dir As String
            dir = ""
            If Len(rawDir) > 0 Then dir = ResolvePath(rawDir)
            ToolProjectSnapshot = ProjShip.SnapshotCreateJson(ResolvePath(rawPath), dir)
        Case "restore"
            If Len(rawDir) = 0 Then Err.Raise vbObjectError + 78, "CodeTools", "dir is required for action=restore"
            Dim csv As String
            csv = ""
            ToolProjectSnapshot = ProjShip.SnapshotRestoreJson(ResolvePath(rawDir), csv)
        Case "list"
            If Len(rawPath) = 0 Then Err.Raise vbObjectError + 78, "CodeTools", "path is required for action=list"
            ToolProjectSnapshot = ProjShip.SnapshotListJson(ResolvePath(rawPath))
        Case Else
            Err.Raise vbObjectError + 78, "CodeTools", "Invalid action: " & action & " (expected create/restore/list)"
    End Select
End Function

Private Function ToolPackageApp(ByVal args As Object) As String
    Dim rawPath As String
    rawPath = JsonGetStr(args, "path")
    If Len(rawPath) = 0 Then Err.Raise vbObjectError + 79, "CodeTools", "path is required (the .vbp to package)"
    Dim exe As String
    exe = JsonGetStr(args, "exe_path")
    If Len(exe) > 0 Then exe = ResolvePath(exe)
    Dim outDir As String
    outDir = JsonGetStr(args, "out_dir")
    If Len(outDir) > 0 Then outDir = ResolvePath(outDir)
    ToolPackageApp = ProjShip.PackageJson(ResolvePath(rawPath), exe, outDir, JsonGetStr(args, "level"))
End Function

Private Function ActiveProjectDir() As String
    Dim f As String
    f = ActiveProjectFile()
    ActiveProjectDir = ""
    If Len(f) = 0 Then Exit Function
    Dim p As Long
    p = InStrRev(f, "\")
    If p <= 1 Then Exit Function
    ActiveProjectDir = Left$(f, p - 1)
End Function

Private Function ToolGit(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim repo As String
    repo = JsonGetStr(args, "repo")
    If Len(repo) = 0 Then repo = ActiveProjectDir()
    If Len(repo) = 0 Then Err.Raise vbObjectError + 80, "CodeTools", "repo is required when no project is open in the IDE"
    repo = ResolvePath(repo)
    Dim gitExe As String
    gitExe = JsonGetStr(args, "exe")
    If Len(gitExe) = 0 Then gitExe = "git.exe"
    Dim sub_ As String
    Dim f As String
    f = JsonGetStr(args, "file")
    Select Case action
        Case "status"
            sub_ = "-C " & Qq(repo) & " status --porcelain=v1 -b"
        Case "diff"
            If Len(f) > 0 Then
                sub_ = "-C " & Qq(repo) & " --no-pager diff --stat -- " & Qq(f)
            Else
                sub_ = "-C " & Qq(repo) & " --no-pager diff --stat"
            End If
        Case "log"
            sub_ = "-C " & Qq(repo) & " --no-pager log --oneline -n " & CStr(JsonGetLong(args, "max", 20))
        Case Else
            Err.Raise vbObjectError + 80, "CodeTools", "Invalid action: " & action & " (expected status/diff/log)"
    End Select
    ToolGit = ProjShip.RunCaptureJson(gitExe, sub_, 30000)
End Function

' ---------------------------------------------------------------------
' project metadata (version resource, app title, conditional compilation)
' ---------------------------------------------------------------------
Private Function ToolProjectMeta(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim vbp As String
    vbp = JsonGetStr(args, "path")
    If Len(vbp) = 0 Then vbp = ActiveProjectFile()
    If Len(vbp) = 0 Then Err.Raise vbObjectError + 65, "CodeTools", "path is required when no project is open in the IDE"
    vbp = ResolvePath(vbp)
    If action = "get" Then
        ToolProjectMeta = ProjMeta.MetaGetJson(vbp)
        Exit Function
    End If
    If action <> "set" Then Err.Raise vbObjectError + 65, "CodeTools", "Invalid action: " & action & " (expected get/set)"
    Dim upd As Object
    Set upd = CreateObject("Scripting.Dictionary")
    PutIfGiven upd, "MajorVer", JsonGetStr(args, "major")
    PutIfGiven upd, "MinorVer", JsonGetStr(args, "minor")
    PutIfGiven upd, "RevisionVer", JsonGetStr(args, "revision")
    If HasArg(args, "auto_increment") Then
        If JsonGetBool(args, "auto_increment", False) Then
            upd("AutoIncrementVer") = "1"
        Else
            upd("AutoIncrementVer") = "0"
        End If
    End If
    PutIfGiven upd, "CompanyName", JsonGetStr(args, "company_name")
    PutIfGiven upd, "FileDescription", JsonGetStr(args, "file_description")
    PutIfGiven upd, "ProductName", JsonGetStr(args, "product_name")
    PutIfGiven upd, "LegalCopyright", JsonGetStr(args, "legal_copyright")
    PutIfGiven upd, "LegalTrademark", JsonGetStr(args, "legal_trademark")
    PutIfGiven upd, "Comments", JsonGetStr(args, "comments")
    PutIfGiven upd, "Title", JsonGetStr(args, "title")
    PutIfGiven upd, "ExeName32", JsonGetStr(args, "exe_name")
    ' TASK-P1-132 (MCP-lead-2 item 4): clearing a macro IS a real operation (a
    ' production build wants the test macros gone), so a PRESENT but EMPTY
    ' cond_comp must be accepted. PutIfGiven skips empty strings, which made
    ' {action:"set", cond_comp:""} raise no-fields-supplied and answer
    ' required_argument_missing. PRESENCE is what selects the write here.
    If HasArg(args, "cond_comp") Then
        upd("CondComp") = JsonGetStr(args, "cond_comp")
    End If
    Dim changed As String
    Dim added As String
    changed = ""
    added = ""
    Dim res As String
    res = ProjMeta.MetaSetJson(vbp, upd, changed, added)
    Dim hint As String
    hint = ""
    If IsOpenProjectFile(vbp) Then hint = "This .vbp belongs to a project that is OPEN in the IDE: the text edit will be lost if the IDE saves the project afterwards. Reopen the project (File > Open) to pick the new values up."
    ToolProjectMeta = WithHint(res, hint)
End Function

' ---------------------------------------------------------------------
' application manifest (UAC level + DPI awareness)
' ---------------------------------------------------------------------
Private Function ToolAppManifest(ByVal args As Object) As String
    Dim action As String
    action = LCase$(JsonGetStr(args, "action"))
    Dim p As String
    p = JsonGetStr(args, "path")
    If Len(p) = 0 Then Err.Raise vbObjectError + 66, "CodeTools", "path is required (the .exe or .dll the manifest belongs to)"
    p = ResolvePath(p)
    If action = "read" Then
        ToolAppManifest = ProjMeta.ManifestReadJson(p)
        Exit Function
    End If
    If action <> "write" Then Err.Raise vbObjectError + 66, "CodeTools", "Invalid action: " & action & " (expected read/write)"
    Dim dpi As Boolean
    dpi = True
    If HasArg(args, "dpi_aware") Then dpi = JsonGetBool(args, "dpi_aware", True)
    ToolAppManifest = ProjMeta.ManifestWriteJson(p, JsonGetStr(args, "level"), dpi, _
        JsonGetStr(args, "dpi_awareness"), JsonGetBool(args, "ui_access", False), JsonGetStr(args, "assembly_name"))
End Function

' ---------------------------------------------------------------------
' reference health (missing reference detection)
' ---------------------------------------------------------------------
Private Function ToolCheckReferences(ByVal args As Object) As String
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim includeResolved As Boolean
    includeResolved = JsonGetBool(args, "include_resolved", False)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim items As Object
    Set items = New Collection
    Dim missingPaths As Object
    Set missingPaths = New Collection
    Dim r As VBIDE.Reference
    Dim total As Long
    Dim missingN As Long
    total = 0
    missingN = 0
    For Each r In p.References
        total = total + 1
        Dim e As Object
        Set e = CreateObject("Scripting.Dictionary")
        Dim fp As String
        On Error Resume Next
        e("name") = r.Name
        e("guid") = r.Guid
        e("major") = r.Major
        e("minor") = r.Minor
        fp = r.FullPath
        On Error GoTo 0
        e("fullPath") = fp
        Dim okResolved As Boolean
        okResolved = True
        ' Built-in references (VBA/VBRUN/VB) are embedded in msvbvm60.dll and
        ' VB6 reports a synthetic path such as "...\msvbvm60.dll\3" for them;
        ' checking that path on disk would always be a false positive.
        Dim isBuiltIn As Boolean
        isBuiltIn = False
        On Error Resume Next
        isBuiltIn = r.BuiltIn
        On Error GoTo 0
        If Not isBuiltIn Then
            If Len(fp) > 0 Then
                On Error Resume Next
                okResolved = fso.FileExists(fp)
                On Error GoTo 0
            End If
        End If
        e("builtIn") = isBuiltIn
        e("resolved") = okResolved
        If okResolved Then
            If includeResolved Then items.Add e
        Else
            missingN = missingN + 1
            items.Add e
            missingPaths.Add fp
        End If
    Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("total") = total
    root("missing") = missingN
    root("ok") = (missingN = 0)
    Set root("references") = items
    Set root("missingPaths") = missingPaths
    ToolCheckReferences = JsonStringify(root)
End Function

' ---- small shared helpers for the tools above ----

Private Function ActiveProjectFile() As String
    ActiveProjectFile = ""
    On Error Resume Next
    If m_vbe Is Nothing Then Exit Function
    If m_vbe.VBProjects.Count = 0 Then Exit Function
    ActiveProjectFile = m_vbe.ActiveVBProject.FileName
End Function

Private Function IsOpenProjectFile(ByVal fullPath As String) As Boolean
    IsOpenProjectFile = False
    On Error Resume Next
    If m_vbe Is Nothing Then Exit Function
    Dim p As VBIDE.VBProject
    For Each p In m_vbe.VBProjects
        If StrComp(p.FileName, fullPath, vbTextCompare) = 0 Then
            IsOpenProjectFile = True
            Exit Function
        End If
    Next
End Function

Private Function HasArg(ByVal args As Object, ByVal key As String) As Boolean
    HasArg = False
    If args Is Nothing Then Exit Function
    On Error Resume Next
    HasArg = args.Exists(key)
End Function

' TASK-P1-112: read vs write for property_bag. The DOCUMENTED trigger is the presence of \`value\`
' ("Value to write (omit to read)"); \`has_value\` is an explicit override for writing an EMPTY value.
' \`value\` can only ADD a write - a documented input must never be silently ignored (discipline 38).
Private Function WantsPropertyWrite(ByVal args As Object) As Boolean
    WantsPropertyWrite = False
    If args Is Nothing Then Exit Function
    If HasArg(args, "value") Then
        WantsPropertyWrite = True
        Exit Function
    End If
    If HasArg(args, "has_value") Then WantsPropertyWrite = JsonGetBool(args, "has_value", False)
End Function

Private Sub PutIfGiven(ByVal d As Object, ByVal key As String, ByVal value As String)
    If Len(value) > 0 Then d(key) = value
End Sub

' Appends a routeHint to an already-serialized JSON object (which ends with }).
Private Function WithHint(ByVal jsonObj As String, ByVal hint As String) As String
    If Len(hint) = 0 Then
        WithHint = jsonObj
        Exit Function
    End If
    WithHint = Left$(jsonObj, Len(jsonObj) - 1) & "," & JsonStringify("routeHint") & ":" & JsonStringify(hint) & "}"
End Function

' ---------------------------------------------------------------------
' test host (T5 / C8-C11)
' TestHost.bas was implemented and verified, but no tool exposed it.
' ---------------------------------------------------------------------
Private Function ToolTestHost(ByVal action As String, ByVal vbpPath As String, ByVal moduleName As String, _
                              ByVal moduleUnderTest As String, ByVal assertionsPath As String, _
                              ByVal runnerPath As String, ByVal exeName As String, ByVal resultPath As String) As String
    Dim a As String
    a = LCase$(action)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("action") = a
    Select Case a
        Case "generate"
            If Len(vbpPath) = 0 Then Err.Raise vbObjectError + 92, "CodeTools", "vbp_path is required for action=generate"
            If Len(moduleUnderTest) = 0 Then Err.Raise vbObjectError + 92, "CodeTools", "module_under_test is required for action=generate"
            If Len(assertionsPath) = 0 Then Err.Raise vbObjectError + 92, "CodeTools", "assertions_path is required for action=generate"
            If Len(runnerPath) = 0 Then Err.Raise vbObjectError + 92, "CodeTools", "runner_path is required for action=generate"
            If Len(exeName) = 0 Then exeName = "vb6testhost.exe"
            If Len(moduleName) = 0 Then moduleName = ModuleNameOf(moduleUnderTest)
            TestHost.GenerateTestHostVbpNamed ResolvePath(vbpPath), moduleName, ResolvePath(moduleUnderTest), _
                                               ResolvePath(assertionsPath), ResolvePath(runnerPath), exeName
            root("vbpPath") = ResolvePath(vbpPath)
            root("moduleName") = moduleName
            root("exeName") = exeName
            root("note") = "Next: build(project=vbp_path), run the produced exe, then action=results with result_path."
        Case "results", "count"
            If Len(resultPath) = 0 Then Err.Raise vbObjectError + 92, "CodeTools", "result_path is required for action=results"
            Dim passN As Long
            Dim failN As Long
            Dim total As Long
            passN = 0
            failN = 0
            total = TestHost.CountResults(ResolvePath(resultPath), passN, failN)
            root("resultPath") = ResolvePath(resultPath)
            root("total") = total
            root("pass") = passN
            root("fail") = failN
            root("allPass") = (failN = 0 And total > 0)
        Case Else
            Err.Raise vbObjectError + 92, "CodeTools", "Invalid action: " & action & " (expected generate/results)"
    End Select
    ToolTestHost = JsonStringify(root)
End Function

' Reads Attribute VB_Name from the head of a .bas/.cls so the generated host
' vbp names the component correctly (the original hard-coded PlanFactLite).
Private Function ModuleNameOf(ByVal path As String) As String
    Dim p As String
    Dim f As Integer
    Dim ln As String
    Dim i As Long
    Dim q1 As Long
    Dim q2 As Long
    ModuleNameOf = ""
    p = ResolvePath(path)
    On Error Resume Next
    f = FreeFile
    Open p For Input As #f
    If Err.Number <> 0 Then
        Err.Clear
        ModuleNameOf = "PlanFactLite"
        On Error GoTo 0
        Exit Function
    End If
    For i = 1 To 10
        If EOF(f) Then Exit For
        Line Input #f, ln
        If InStr(1, ln, "Attribute VB_Name", vbTextCompare) > 0 Then
            q1 = InStr(ln, Chr$(34))
            If q1 > 0 Then
                q2 = InStr(q1 + 1, ln, Chr$(34))
                If q2 > q1 Then ModuleNameOf = Mid$(ln, q1 + 1, q2 - q1 - 1)
            End If
            Exit For
        End If
    Next
    Close #f
    On Error GoTo 0
    If Len(ModuleNameOf) = 0 Then ModuleNameOf = "PlanFactLite"
End Function

' Card TASK-P1-005: keep the 2-arg form for the native-breakpoint caller, and add a
' variant that reports WHY nothing was executed.
Private Function ExecIfEnabled(ByVal ctrl As Object, ByVal m As String) As Long
    Dim w As String
    ExecIfEnabled = ExecIfEnabledWhy(ctrl, m, w)
End Function

Private Function ExecIfEnabledWhy(ByVal ctrl As Object, ByVal m As String, ByRef whyOut As String) As Long
    Dim ok As Boolean
    whyOut = ""
    If ctrl Is Nothing Then
        whyOut = "control_missing"
        Exit Function
    End If
    ok = False
    On Error Resume Next
    ok = ctrl.Enabled
    If Err.Number <> 0 Then whyOut = "enabled_query_failed"
    Err.Clear
    On Error GoTo 0
    If ok Then
        ctrl.Execute
        ExecIfEnabledWhy = 1
    ElseIf Len(whyOut) = 0 Then
        whyOut = "item_disabled"
    End If
End Function

' Card TASK-P1-005 fallback 2: FindBarControl (Util.bas) filters bars by NAME and
' descends only ONE popup level. In break state the bar hosting the step items
' answers to a name outside the candidate list, so walk EVERY CommandBar
' recursively by caption before giving up.
Private Function FindBarControlDeep(ByVal ctrlCands As String) As Object
    On Error Resume Next
    Dim bars As Object, b As Object, i As Long
    Set bars = m_vbe.CommandBars
    If bars Is Nothing Then Exit Function
    For i = 1 To bars.Count
        Set b = bars.Item(i)
        If Not b Is Nothing Then
            Set FindBarControlDeep = FindControlDeep(b.Controls, ctrlCands, 0)
            If Not FindBarControlDeep Is Nothing Then Exit Function
        End If
    Next
End Function

Private Function FindControlDeep(ByVal col As Object, ByVal ctrlCands As String, ByVal depth As Long) As Object
    On Error Resume Next
    If col Is Nothing Then Exit Function
    If depth > 3 Then Exit Function
    Dim n As Long, k As Long, c As Object, subCol As Object
    n = col.Count
    For k = 1 To n
        Set c = col.Item(k)
        If Not c Is Nothing Then
            If MatchCand(c.Caption, ctrlCands) Then
                Set FindControlDeep = c
                Exit Function
            End If
            ' a button has no children (touching .Controls raises 438). The Set must
            ' be RESET first, or a failed access leaves the PREVIOUS level in subCol
            ' and the recursion walks the wrong collection.
            Set subCol = Nothing
            Err.Clear
            Set subCol = c.Controls
            If Not subCol Is Nothing Then
                Set FindControlDeep = FindControlDeep(subCol, ctrlCands, depth + 1)
                If Not FindControlDeep Is Nothing Then Exit Function
            End If
        End If
    Next
End Function

Public Sub DoDeferredStart()
    ' called by the server watchdog after a deferred restart: the IDE
    ' is back in design state, so the Start button is enabled again
    On Error Resume Next
    Dim ctrl As Object
    Set ctrl = FindCmd("start", "Debug|Run|" & ChrW$(&H8C03) & ChrW$(&H8BD5) & "|" & ChrW$(&H8FD0) & ChrW$(&H884C), "Start|" & ChrW$(&H542F) & ChrW$(&H52A8) & "|Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED))
    If ctrl Is Nothing Then Set ctrl = m_cachedStart Else Set m_cachedStart = ctrl
    If Not ctrl Is Nothing Then
        Dim ok As Boolean
        ok = False
        ok = ctrl.Enabled
        If ok Then ctrl.Execute
    End If
End Sub

' =====================================================================
' reference tools
' =====================================================================

Private Function ToolAddReference(ByVal mode As String, ByVal guid As String, ByVal path As String, ByVal major As Long, ByVal minor As Long) As String
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim ref As VBIDE.Reference
    If LCase$(mode) = "file" Then
        If Len(path) = 0 Then Err.Raise vbObjectError + 55, "CodeTools", "path is required for mode=file"
        Set ref = p.References.AddFromFile(ResolvePath(path))
    Else
        If Len(guid) = 0 Then Err.Raise vbObjectError + 42, "CodeTools", "guid is required for mode=guid"
        Set ref = p.References.AddFromGuid(guid, major, minor)
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("name") = ref.Name
    root("description") = ref.Description
    root("guid") = ref.Guid
    root("major") = ref.Major
    root("minor") = ref.Minor
    root("fullPath") = ref.FullPath
    root("type") = RefKindName(ref.Type)
    ToolAddReference = JsonStringify(root)
End Function

Private Function ToolRemoveReference(ByVal name As String) As String
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim ref As VBIDE.Reference
    Dim target As VBIDE.Reference
    Dim refName As String
    Dim foundRef As Boolean
    foundRef = False
    For Each ref In p.References
        If LCase$(ref.Name) = LCase$(name) Or LCase$(ref.Guid) = LCase$(name) Then
            Set target = ref
            refName = ref.Name
            foundRef = True
            Exit For
        End If
    Next
    If Not foundRef Then Err.Raise vbObjectError + 56, "CodeTools", "Reference not found: " & name
    On Error Resume Next
    p.References.Remove target
    Dim rmErr As Long
    Dim rmDesc As String
    rmErr = Err.Number
    rmDesc = Err.Description
    On Error GoTo 0
    If rmErr <> 0 Then Err.Raise vbObjectError + 56, "CodeTools", "References.Remove failed: " & rmDesc
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("removed") = refName
    ToolRemoveReference = JsonStringify(root)
End Function

' =====================================================================
' design-time property tools
' =====================================================================

Private Function ToolGetComponentProperties(ByVal name As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim arr As Object
    Set arr = New Collection
    On Error Resume Next
    Dim props As Object
    Set props = comp.Properties
    If props Is Nothing Then
        On Error GoTo 0
        Err.Raise vbObjectError + 57, "CodeTools", "Component has no Properties collection (not a designer component): " & name
    End If
    Dim prop As Object
    For Each prop In props
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = prop.Name
        d("numIndices") = prop.NumIndices
        Err.Clear
        Dim v As Variant
        Dim vObj As Object
        If IsObject(prop.Value) Then
            Set vObj = prop.Value
            d("value") = TypeName(vObj)
        Else
            v = prop.Value
            If Err.Number = 0 Then
                If IsNull(v) Then
                    d("value") = "(Null)"
                Else
                    d("value") = CStr(v)
                End If
            Else
                d("value") = "(unreadable)"
            End If
        End If
        arr.Add d
    Next
    On Error GoTo 0
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("count") = arr.Count
    Set root("properties") = arr
    ToolGetComponentProperties = JsonStringify(root)
End Function

Private Function ToolSetComponentProperty(ByVal name As String, ByVal propName As String, ByVal value As String, ByVal indices As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim props As Object
    Set props = comp.Properties
    If props Is Nothing Then Err.Raise vbObjectError + 57, "CodeTools", "Component has no Properties collection: " & name
    Dim prop As Object
    Dim found As Boolean
    found = False
    For Each prop In props
        If LCase$(prop.Name) = LCase$(propName) Then
            found = True
            Exit For
        End If
    Next
    If Not found Then Err.Raise vbObjectError + 58, "CodeTools", "Property not found: " & propName
    Dim parsed As Variant
    parsed = ParseVariantValue(value)
    If Len(indices) > 0 Then
        Dim idxParts() As String
        idxParts = Split(indices, ",")
        Select Case UBound(idxParts)
            Case 0
                prop.IndexedValue(CLng(Val(idxParts(0)))) = parsed
            Case 1
                prop.IndexedValue(CLng(Val(idxParts(0))), CLng(Val(idxParts(1)))) = parsed
            Case 2
                prop.IndexedValue(CLng(Val(idxParts(0))), CLng(Val(idxParts(1))), CLng(Val(idxParts(2)))) = parsed
            Case Else
                prop.IndexedValue(CLng(Val(idxParts(0))), CLng(Val(idxParts(1))), CLng(Val(idxParts(2))), CLng(Val(idxParts(3)))) = parsed
        End Select
    Else
        prop.Value = parsed
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("property") = prop.Name
    root("value") = value
    ToolSetComponentProperty = JsonStringify(root)
End Function

' Accepts "0x1F2E", "1F2Eh", "&#x1F2E" or plain decimal - used by vba6_probe rva_call.
Private Function ParseHexOrLong(ByVal s As String) As Long
    Dim t As String
    Dim i As Long
    Dim ch As String
    Dim v As Long
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseHexOrLong = 0
        Exit Function
    End If
    If LCase$(Left$(t, 2)) = "0x" Then t = Mid$(t, 3)
    If LCase$(Right$(t, 1)) = "h" Then t = Left$(t, Len(t) - 1)
    If LCase$(Left$(t, 3)) = "&#x" Then t = Mid$(t, 4)
    v = 0
    For i = 1 To Len(t)
        ch = LCase$(Mid$(t, i, 1))
        If ch >= "0" And ch <= "9" Then
            v = v * 16 + (Asc(ch) - Asc("0"))
        ElseIf ch >= "a" And ch <= "f" Then
            v = v * 16 + (Asc(ch) - Asc("a") + 10)
        End If
    Next
    ParseHexOrLong = v
End Function

Private Function ParseVariantValue(ByVal s As String) As Variant
    Dim t As String
    t = Trim$(s)
    If LCase$(t) = "true" Then
        ParseVariantValue = True
    ElseIf LCase$(t) = "false" Then
        ParseVariantValue = False
    ElseIf LCase$(t) = "null" Or LCase$(t) = "(nothing)" Then
        ParseVariantValue = Null
    ElseIf IsNumeric(t) Then
        If InStr(t, ".") > 0 Or InStr(1, t, "e", vbTextCompare) > 0 Then
            ParseVariantValue = Val(t)
        Else
            ParseVariantValue = CLng(Val(t))
        End If
    Else
        ParseVariantValue = s
    End If
End Function

' =====================================================================
' member tools
' =====================================================================

Private Function ToolListMembers(ByVal component As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(component)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
    Dim arr As Object
    Set arr = New Collection
    On Error Resume Next
    Dim m As VBIDE.Member
    For Each m In cm.Members
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = m.Name
        d("type") = MemberTypeName(m.Type)
        d("scope") = ScopeName(m.Scope)
        d("static") = m.Static
        d("codeLocation") = m.CodeLocation
        d("description") = m.Description
        d("hidden") = m.Hidden
        d("browsable") = m.Browsable
        d("uiDefault") = m.UIDefault
        d("category") = m.Category
        d("propertyPage") = m.PropertyPage
        arr.Add d
    Next
    On Error GoTo 0
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("count") = arr.Count
    Set root("members") = arr
    ToolListMembers = JsonStringify(root)
End Function

Private Function ToolSetMemberAttribute(ByVal component As String, ByVal member As String, ByVal attrName As String, ByVal value As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(component)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
    Dim m As VBIDE.Member
    Dim found As Boolean
    found = False
    For Each m In cm.Members
        If LCase$(m.Name) = LCase$(member) Then
            found = True
            Exit For
        End If
    Next
    If Not found Then Err.Raise vbObjectError + 59, "CodeTools", "Member not found: " & member
    Dim b As Boolean
    b = (LCase$(value) = "true" Or value = "1")
    Select Case LCase$(attrName)
        Case "description"
            m.Description = value
        Case "help_context_id"
            m.HelpContextID = CLng(Val(value))
        Case "hidden"
            m.Hidden = b
        Case "browsable"
            m.Browsable = b
        Case "ui_default"
            m.UIDefault = b
        Case "bindable"
            m.Bindable = b
        Case "default_bind"
            m.DefaultBind = b
        Case "display_bind"
            m.DisplayBind = b
        Case "request_edit"
            m.RequestEdit = b
        Case "category"
            m.Category = value
        Case "property_page"
            m.PropertyPage = value
        Case "standard_method"
            m.StandardMethod = CLng(Val(value))
        Case Else
            Err.Raise vbObjectError + 60, "CodeTools", "Unknown attrName: " & attrName
    End Select
    ' TASK-P1-164 (R18-2): a write-class tool reads its own write BACK, so a setter that
    ' silently did nothing cannot be mistaken for a success.
    Dim rb As String
    Dim want As String
    rb = ""
    want = value
    Select Case LCase$(attrName)
        Case "description"
            rb = m.Description
        Case "help_context_id"
            rb = CStr(m.HelpContextID)
            want = CStr(CLng(Val(value)))
        Case "hidden"
            rb = CStr(m.Hidden)
            want = CStr(b)
        Case "browsable"
            rb = CStr(m.Browsable)
            want = CStr(b)
        Case "ui_default"
            rb = CStr(m.UIDefault)
            want = CStr(b)
        Case "bindable"
            rb = CStr(m.Bindable)
            want = CStr(b)
        Case "default_bind"
            rb = CStr(m.DefaultBind)
            want = CStr(b)
        Case "display_bind"
            rb = CStr(m.DisplayBind)
            want = CStr(b)
        Case "request_edit"
            rb = CStr(m.RequestEdit)
            want = CStr(b)
        Case "category"
            rb = m.Category
        Case "property_page"
            rb = m.PropertyPage
        Case "standard_method"
            rb = CStr(m.StandardMethod)
            want = CStr(CLng(Val(value)))
    End Select
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = comp.Name
    root("member") = m.Name
    root("attribute") = attrName
    root("requestedValue") = value
    root("readback") = rb
    root("applied") = (LCase$(Trim$(rb)) = LCase$(Trim$(want)))
    root("readbackNote") = "readback is the value the IDE reports for this member attribute AFTER the set; applied compares it with the normalised request. It is the IN-MEMORY value only - whether it reaches the FILE on disk is up to VB6 when the component is saved (save_project), and member attributes such as hidden/uidDefault are written as Attribute lines the code pane never shows."
    ToolSetMemberAttribute = JsonStringify(root)
End Function

' =====================================================================
' window tools
' =====================================================================

Private Function ToolListWindows() As String
    Dim arr As Object
    Set arr = New Collection
    Dim w As VBIDE.Window
    For Each w In m_vbe.Windows
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("caption") = w.Caption
        d("type") = WindowTypeName(w.Type)
        d("visible") = w.Visible
        d("windowState") = WindowStateName(w.WindowState)
        d("left") = w.Left
        d("top") = w.Top
        d("width") = w.Width
        d("height") = w.Height
        d("hwnd") = w.HWnd
        arr.Add d
    Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("count") = arr.Count
    Set root("windows") = arr
    ToolListWindows = JsonStringify(root)
End Function

Private Function ToolControlWindow(ByVal window As String, ByVal action As String, ByVal caption As String, ByVal left As Long, ByVal top As Long, ByVal width As Long, ByVal height As Long, ByVal captionGiven As Boolean) As String
    Dim w As VBIDE.Window
    Set w = FindWindow(window)
    If w Is Nothing Then
        ' FindWindow used to swallow its own Err.Raise (the whole body sat under
        ' On Error Resume Next), so it returned Nothing and the next line died with
        ' VB 91 "Object variable or With block variable not set". Report a usable
        ' failure and list what IS addressable.
        Dim miss As Object
        Set miss = CreateObject("Scripting.Dictionary")
        miss("ok") = False
        miss("reason") = "window_not_found"
        miss("tool") = "control_window"
        miss("argument") = "window"
        miss("value") = window
        miss("error") = "window not found: " & window
        miss("suggestedAction") = "call list_windows and use a type name, caption substring or 1-based index it can address (see 'available')"
        miss("sideEffects") = "none: the call was refused before any window was changed"
        miss("hint") = "accepts a locale-independent TYPE name (CodeWindow, ProjectWindow, Immediate, Locals, Watch, Toolbox, ObjectBrowser, PropertyWindow, Designer, Find, Preview, ColorPalette, MainWindow, LinkedWindowFrame), a caption substring, or a 1-based index"
        Set miss("available") = WindowsBriefJson()
        ToolControlWindow = JsonStringify(miss)
        Exit Function
    End If
    ' TASK-P1-164 (R18 focus safety): record the foreground window across EVERY action so the
    ' receipt can show whether driving the IDE stole the user's foreground. An action that did
    ' NOT ask to activate (move/resize/show/...) gives the foreground BACK - VB6 can raise the
    ' frame as a side effect, and that is exactly how a user's typing lands in the wrong window.
    Dim fgBefore As Long
    Dim fgAfter As Long
    ' TASK-P1-164 (A): set when action=activate skipped Window.SetFocus because quiet mode is ON.
    Dim quietSuppressed As Boolean
    quietSuppressed = False
    ' TASK-P1-164 (R18-2): a window action writes no text, so the contract is "the pane
    ' text is UNCHANGED". When the addressed window IS a code window, its component is
    ' identified by the WINDOW HANDLE (not by a caption guess) and its pane is compared.
    Dim pAssertComp As String, pAssertBefore As String
    Dim pAssertOn As Boolean
    Dim pAssertVia As String
    pAssertComp = ""
    pAssertBefore = ""
    pAssertOn = False
    pAssertVia = ""
    pAssertComp = PaneCompOfWindowTarget(w.HWnd, w.Caption, pAssertVia)
    If Len(pAssertComp) > 0 Then pAssertBefore = PaneTextOfName(pAssertComp, pAssertOn)
    fgBefore = WoFrontHwnd()
    Select Case LCase$(action)
        Case "activate"
            ' TASK-P1-164 (A): the ONE place where the add-in deliberately takes the user's
            ' foreground. With quiet mode ON the SetFocus is NOT issued - the window is still
            ' resolved and every other part of the action is performed.
            If QuietActive() Then
                quietSuppressed = True
                m_quietSuppressed = m_quietSuppressed + 1
            Else
                w.SetFocus
            End If
        Case "close"
            w.Close
        Case "show"
            w.Visible = True
        Case "hide"
            w.Visible = False
        Case "minimize"
            w.WindowState = vbext_ws_Minimize
        Case "maximize"
            w.WindowState = vbext_ws_Maximize
        Case "restore"
            w.WindowState = vbext_ws_Normal
        Case "move"
            If left <> 0 Then w.Left = left
            If top <> 0 Then w.Top = top
        Case "resize"
            If width <> 0 Then w.Width = width
            If height <> 0 Then w.Height = height
        Case "set_caption"
            ' Window.Caption is read-only in VB6EXT; acknowledged but ignored.
        Case Else
            Err.Raise vbObjectError + 61, "CodeTools", "Unknown action: " & action
    End Select
    fgAfter = WoFrontHwnd()
    If fgAfter <> fgBefore Then
        If LCase$(action) <> "activate" Then
            ' The action did not request the foreground, so put the user's window back and
            ' re-read: the receipt reports what is true AFTER the restore attempt.
            WoFrontRestore fgBefore
            fgAfter = WoFrontHwnd()
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("ok") = True
    root("foregroundBefore") = fgBefore
    root("foregroundAfter") = fgAfter
    root("foregroundTookFocus") = (fgAfter <> fgBefore)
    root("caption") = w.Caption
    root("visible") = w.Visible
    root("windowState") = WindowStateName(w.WindowState)
    root("type") = WindowTypeName(w.Type)
    ' TASK-P1-114: VB6EXT v5.3 declares Window.Caption GET-ONLY (the Window put-map holds
    ' only Height/Left/Top/Visible/Width/WindowState), so \`caption\` can NEVER be applied.
    ' Report it instead of accepting it silently (discipline 38). Gated on the argument
    ' actually being present, so a call that does not carry it stays byte-identical.
    If captionGiven Then
        root("ignoredArgs") = Array("caption")
        root("ignoredArgsNote") = "VB6's extensibility model exposes Window.Caption as READ-ONLY (VB6EXT v5.3), so " & _
            "\`caption\` cannot be applied: action=set_caption is accepted but the window is NOT renamed. The window's " & _
            "CURRENT caption is in this receipt. The action you did request was performed (see visible/windowState/type); " & _
            "use move / resize / show / hide / minimize / maximize / restore to change the window instead."
    End If
    ' R18-2: see E1 - a window action must not touch the code text.
    PaneAssertFinish root, pAssertComp, pAssertBefore, pAssertOn, "control_window"
    root("paneAssertionVia") = pAssertVia
    ' TASK-P1-164 (R18-2) honesty: a foreground handle of 0 means the foreground could not
    ' be READ in this session station - it does NOT mean the operation left the focus alone.
    root("foregroundReadable") = (fgBefore <> 0 Or fgAfter <> 0)
    ' TASK-P1-164 (A): report the suppression ON THE RECEIPT that caused it. The key only
    ' exists when it actually happened, so a quiet-OFF call stays byte-identical.
    If quietSuppressed Then
        root("quietModeSuppressedFocus") = True
        root("suppressedAction") = "control_window{action:""activate""}: Window.SetFocus was NOT called because ide_quiet_mode is ON"
        root("quietModeNote") = "quiet mode is ON: this action did NOT ask for the foreground. The window was resolved and every other part of the action was performed. NOTE: foregroundBefore/foregroundAfter can both be 0 in this window station - foregroundReadable=false says the handle could not be READ, not that the focus was left alone."
    End If
    If Len(pAssertComp) = 0 Then
        root("paneAssertionTarget") = "no code pane could be matched to the addressed window (tried hwnd identity, then exact caption equality), so the assertion is reported as notApplicable - never as a pass"
    End If
    ToolControlWindow = JsonStringify(root)
End Function
Private Function WindowTypeOfName(ByVal s As String) As Long
    Dim k As String
    WindowTypeOfName = -1
    k = LCase$(Trim$(s))
    If Len(k) = 0 Then Exit Function
    Select Case k
        Case "codewindow"
            WindowTypeOfName = vbext_wt_CodeWindow
        Case "designer"
            WindowTypeOfName = vbext_wt_Designer
        Case "browser", "objectbrowser"
            WindowTypeOfName = vbext_wt_Browser
        Case "watch"
            WindowTypeOfName = vbext_wt_Watch
        Case "locals"
            WindowTypeOfName = vbext_wt_Locals
        Case "immediate"
            WindowTypeOfName = vbext_wt_Immediate
        Case "projectwindow", "project"
            WindowTypeOfName = vbext_wt_ProjectWindow
        Case "propertywindow", "properties"
            WindowTypeOfName = vbext_wt_PropertyWindow
        Case "find"
            WindowTypeOfName = vbext_wt_Find
        Case "findreplace"
            WindowTypeOfName = vbext_wt_FindReplace
        Case "toolbox"
            WindowTypeOfName = vbext_wt_Toolbox
        Case "linkedwindowframe"
            WindowTypeOfName = vbext_wt_LinkedWindowFrame
        Case "mainwindow"
            WindowTypeOfName = vbext_wt_MainWindow
        Case "preview"
            WindowTypeOfName = vbext_wt_Preview
        Case "colorpalette"
            WindowTypeOfName = vbext_wt_ColorPalette
        Case Else
            WindowTypeOfName = -1
    End Select
End Function

' Snapshot of the addressable windows, for the "not found" reply.
Private Function WindowsBriefJson() As Object
    Dim arr As Object, w As VBIDE.Window, d As Object
    Set arr = New Collection
    On Error Resume Next
    For Each w In m_vbe.Windows
        Set d = CreateObject("Scripting.Dictionary")
        d("index") = arr.Count + 1
        d("type") = WindowTypeName(w.Type)
        d("caption") = w.Caption
        d("visible") = w.Visible
        arr.Add d
    Next
    Err.Clear
    On Error GoTo 0
    Set WindowsBriefJson = arr
End Function

' Three lookup strategies, each with its OWN scoped error handling. The old body
' put On Error Resume Next over everything and ended with Err.Raise, which is
' silently swallowed under Resume Next -> it returned Nothing instead.
Private Function FindWindow(ByVal window As String) As VBIDE.Window
    Dim w As VBIDE.Window, idx As Long, want As Long
    If Len(Trim$(window)) = 0 Then Exit Function
    If IsNumeric(window) Then
        On Error Resume Next
        idx = CLng(Val(window))
        Set w = m_vbe.Windows.Item(idx)
        Err.Clear
        On Error GoTo 0
        If Not w Is Nothing Then
            Set FindWindow = w
            Exit Function
        End If
    End If
    want = WindowTypeOfName(window)
    If want >= 0 Then
        On Error Resume Next
        For Each w In m_vbe.Windows
            If CLng(w.Type) = want Then
                Set FindWindow = w
                Exit For
            End If
        Next
        Err.Clear
        On Error GoTo 0
        If Not FindWindow Is Nothing Then Exit Function
    End If
    On Error Resume Next
    For Each w In m_vbe.Windows
        If InStr(1, w.Caption, window, vbTextCompare) > 0 Then
            Set FindWindow = w
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Function

' =====================================================================
' form designer tools
' =====================================================================

Private Function ToolFormDesigner(ByVal form As String, ByVal action As String, ByVal control As String, ByVal index As Long, ByVal parent As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(form)
    Dim des As Object
    Set des = comp.Designer
    If des Is Nothing Then Err.Raise vbObjectError + 63, "CodeTools", "Component has no designer: " & form
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("form") = comp.Name
    Select Case LCase$(action)
        Case "info"
            root("canPaste") = des.CanPaste
            Dim arr As Object
            Set arr = New Collection
            On Error Resume Next
            Dim ctl As Object
            For Each ctl In des.VBControls
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("className") = ctl.ClassName
                d("progId") = ctl.ProgId
                d("inSelection") = ctl.InSelection
                ' 1-based index so set_control_property/remove_control can be
                ' driven by name first and index as the fallback
                d("index") = arr.Count + 1
                d("name") = ControlName(ctl)
                Err.Clear
                d("controlType") = ctl.ControlType
                If Err.Number <> 0 Then d("controlType") = "?"
                arr.Add d
            Next
            On Error GoTo 0
            root("controlCount") = arr.Count
            Set root("controls") = arr
        Case "props"
            Dim prpIdx As Long
            prpIdx = ResolveControlIndex(des, control, index)
            Dim prpCtl As Object
            Set prpCtl = GetControlByIndex(des, prpIdx)
            root("index") = prpIdx
            root("control") = ControlName(prpCtl)
            Err.Clear
            root("controlClass") = prpCtl.ClassName
            If Err.Number <> 0 Then
                Err.Clear
                root("controlClass") = ""
            End If
            Dim prpBag As Object
            Set prpBag = CreateObject("Scripting.Dictionary")
            Dim prpIt As Object
            Dim prpKey As String
            Dim prpVal As String
            On Error Resume Next
            For Each prpIt In prpCtl.Properties
                prpKey = ""
                Err.Clear
                prpKey = prpIt.Name
                If Err.Number <> 0 Then
                    Err.Clear
                    prpKey = ""
                End If
                If Len(prpKey) > 0 Then
                    prpVal = ""
                    Err.Clear
                    prpVal = CStr(prpIt.Value)
                    If Err.Number <> 0 Then
                        Err.Clear
                        prpVal = "<object>"
                    End If
                    prpBag(prpKey) = prpVal
                End If
            Next
            On Error GoTo 0
            Set root("properties") = prpBag
        Case "menu_nest"
            ' VB6 exposes no menu-level property, so the menu tree only exists as
            ' the indentation of the Begin VB.Menu blocks in the .frm: flush the
            ' form to disk, move the block inside the parent, reload.
            ToolFormDesigner = NestMenuItem(comp, control, parent)
            Exit Function
        Case "select_all"
            des.SelectAll
            root("selectedAll") = True
        Case "clear_selection"
            des.SelectedVBControls.Clear
            root("cleared") = True
        Case "cut"
            des.SelectedVBControls.Cut
            root("cut") = True
        Case "copy"
            des.SelectedVBControls.Copy
            root("copied") = True
        Case "paste"
            If Not des.CanPaste Then Err.Raise vbObjectError + 64, "CodeTools", "Clipboard has nothing pasteable for this form"
            des.Paste
            root("pasted") = True
        Case Else
            Err.Raise vbObjectError + 65, "CodeTools", "Unknown action: " & action
    End Select
    ToolFormDesigner = JsonStringify(root)
End Function

' ---- form menu nesting ---------------------------------------------------
' VB6 exposes NO menu-level property: a Menu control's design-time properties
' are only Name/Index/Visible/Caption/Checked/Enabled/Shortcut/Tag/WindowList/
' HelpContextID/NegotiatePosition, and the tree exists solely as the
' indentation of the Begin VB.Menu blocks in the .frm. Measured on a real IDE:
' moving a child block inside its parent's Begin/End, re-indenting it and
' calling VBComponent.Reload makes the IDE adopt the nesting and re-emit the
' file canonically (the IDE also creates the *_Click event stubs itself).
Private Function NestMenuItem(ByVal comp As VBIDE.VBComponent, ByVal child As String, ByVal parent As String) As String
    If Len(child) = 0 Then Err.Raise vbObjectError + 69, "CodeTools", "control (the VB.Menu item to move) is required"
    Dim path As String
    path = ""
    On Error Resume Next
    path = comp.FileNames(1)
    If Err.Number <> 0 Then
        Err.Clear
        path = ""
    End If
    On Error GoTo 0
    If Len(path) = 0 Then Err.Raise vbObjectError + 69, "CodeTools", "the form has no file on disk yet; run save_project first"
    ' flush the live designer state, otherwise Reload would silently discard it
    comp.SaveAs path
    Dim lines() As String
    If Not ReadLinesCp(path, lines) Then Err.Raise vbObjectError + 69, "CodeTools", "could not read " & path
    Dim c1 As Long, c2 As Long, cInd As Long
    If Not FindMenuBlockLines(lines, child, c1, c2, cInd) Then Err.Raise vbObjectError + 69, "CodeTools", "menu item not found in the .frm text: " & child
    Dim p1 As Long, p2 As Long, pInd As Long
    If Len(parent) > 0 Then
        If Not FindMenuBlockLines(lines, parent, p1, p2, pInd) Then Err.Raise vbObjectError + 69, "CodeTools", "parent menu item not found in the .frm text: " & parent
        If c1 > p1 And c1 < p2 Then Err.Raise vbObjectError + 69, "CodeTools", "already a child of " & parent
        If p1 > c1 And p1 < c2 Then Err.Raise vbObjectError + 69, "CodeTools", "cannot nest a menu item inside itself"
    End If
    Dim nblk As Long
    nblk = c2 - c1 + 1
    Dim blk() As String
    ReDim blk(nblk - 1)
    Dim i As Long
    For i = 0 To nblk - 1
        blk(i) = lines(c1 + i)
    Next
    Dim kept() As String
    ReDim kept(UBound(lines) - nblk)
    Dim k As Long
    k = 0
    For i = 0 To UBound(lines)
        If i < c1 Or i > c2 Then
            kept(k) = lines(i)
            k = k + 1
        End If
    Next
    Dim targetIndent As Long
    Dim insertAt As Long
    If Len(parent) > 0 Then
        Dim pp2 As Long
        pp2 = p2
        If p1 > c2 Then pp2 = p2 - nblk
        targetIndent = pInd + 3
        insertAt = pp2
    Else
        targetIndent = 3
        Dim stopAt As Long
        stopAt = UBound(kept)
        For i = 0 To UBound(kept)
            If Left$(Trim$(kept(i)), 10) = "Attribute " Then
                stopAt = i - 1
                Exit For
            End If
        Next
        insertAt = -1
        For i = stopAt To 0 Step -1
            If Trim$(kept(i)) = "End" Then
                insertAt = i
                Exit For
            End If
        Next
        If insertAt < 0 Then Err.Raise vbObjectError + 69, "CodeTools", "could not find the form's closing End line"
    End If
    For i = 0 To nblk - 1
        If Len(Trim$(blk(i))) = 0 Then
            blk(i) = ""
        Else
            blk(i) = Space$(targetIndent) & Mid$(blk(i), cInd + 1)
        End If
    Next
    Dim outp() As String
    ReDim outp(UBound(kept) + nblk)
    k = 0
    For i = 0 To insertAt - 1
        outp(k) = kept(i)
        k = k + 1
    Next
    For i = 0 To nblk - 1
        outp(k) = blk(i)
        k = k + 1
    Next
    For i = insertAt To UBound(kept)
        outp(k) = kept(i)
        k = k + 1
    Next
    Dim bak As String
    bak = path & ".bak"
    On Error Resume Next
    Kill bak
    FileCopy path, bak
    On Error GoTo 0
    WriteLinesCp path, outp
    comp.Reload
    ' verify the nesting really landed instead of trusting the write
    Dim vl() As String
    If Not ReadLinesCp(path, vl) Then Err.Raise vbObjectError + 69, "CodeTools", "could not re-read " & path
    Dim v1 As Long, v2 As Long, vInd As Long
    Dim ok As Boolean
    ok = False
    If FindMenuBlockLines(vl, child, v1, v2, vInd) Then
        If Len(parent) > 0 Then
            Dim q1 As Long, q2 As Long, qInd As Long
            If FindMenuBlockLines(vl, parent, q1, q2, qInd) Then
                If v1 > q1 And v1 < q2 Then ok = True
            End If
        ElseIf vInd = 3 Then
            ok = True
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("form") = comp.Name
    root("control") = child
    root("parent") = parent
    root("file") = path
    root("backup") = bak
    root("verified") = ok
    root("note") = "VB6 has no menu-level property: the tree lives in the .frm indentation, so the form was flushed to disk, the menu block was re-nested, and the component was reloaded. The IDE creates the *_Click event stubs for menu items itself."
    NestMenuItem = JsonStringify(root)
End Function

Private Function FindMenuBlockLines(ByRef lines() As String, ByVal menuName As String, ByRef b1 As Long, ByRef b2 As Long, ByRef ind As Long) As Boolean
    Dim i As Long, j As Long, sp As Long, depth As Long
    Dim t As String, rest As String
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Left$(t, 14) = "Begin VB.Menu " Then
            rest = Trim$(Mid$(t, 15))
            sp = InStr(rest, " ")
            If sp > 0 Then rest = Left$(rest, sp - 1)
            If StrComp(rest, menuName, vbTextCompare) = 0 Then
                ind = Len(lines(i)) - Len(LTrim$(lines(i)))
                depth = 0
                For j = i To UBound(lines)
                    t = Trim$(lines(j))
                    If Left$(t, 6) = "Begin " Then
                        depth = depth + 1
                    ElseIf t = "End" Then
                        depth = depth - 1
                        If depth = 0 Then
                            b1 = i
                            b2 = j
                            FindMenuBlockLines = True
                            Exit Function
                        End If
                    End If
                Next
                Exit Function
            End If
        End If
    Next
End Function

' VB6 writes .frm/.bas in the system ANSI code page, and its own text-mode
' Input/Output statements convert to and from Unicode correctly, so line-based
' reads/writes round-trip the file (a UTF-8 read here would mojibake it).
' Line Input is used on purpose: "Input$(LOF(f), f)" over a text-mode handle
' raises error 62 (input past end of file) on this build.
Private Function ReadLinesCp(ByVal path As String, ByRef lines() As String) As Boolean
    Dim fnum As Integer
    Dim s As String
    Dim n As Long
    Dim buf() As String
    n = 0
    ReDim buf(255)
    fnum = FreeFile
    On Error GoTo Fail
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, s
        If n > UBound(buf) Then ReDim Preserve buf(UBound(buf) * 2 + 1)
        buf(n) = s
        n = n + 1
    Loop
    Close #fnum
    If n = 0 Then
        ReDim lines(0)
        lines(0) = ""
    Else
        ReDim Preserve buf(n - 1)
        lines = buf
    End If
    ReadLinesCp = True
    Exit Function
Fail:
    On Error Resume Next
    Close #fnum
    On Error GoTo 0
    ReadLinesCp = False
End Function

Private Sub WriteLinesCp(ByVal path As String, ByRef lines() As String)
    Dim fnum As Integer
    Dim i As Long
    fnum = FreeFile
    Open path For Output As #fnum
    For i = 0 To UBound(lines)
        MaybePump
        Print #fnum, lines(i)
    Next
    Close #fnum
End Sub

' Picture/Icon/... properties hold a StdPicture, so a file path is loaded
' instead of being parsed as a scalar; the designer serialises it into the .frx
' the next time the form is saved.
' The designer exposes a live control instance through VBControl.Object, whose
' properties are the real runtime ones, so a picture can be pushed through it
' (assigning to the designer property itself can raise 425).
' Returns "" when applied, otherwise a description of which step failed.
Private Function TrySetPicture(ByVal ctl As Object, ByVal propName As String, ByVal pic As StdPicture) As String
    Dim o As Object
    Dim rs As String
    rs = ""
    On Error Resume Next
    Set o = ctl.Object
    If Err.Number <> 0 Then
        rs = "VBControl.Object raised " & CStr(Err.Number) & " " & Err.Description
        Err.Clear
        On Error GoTo 0
        TrySetPicture = rs
        Exit Function
    End If
    Err.Clear
    Select Case LCase$(propName)
        Case "picture", "image"
            Set o.Picture = pic
        Case "icon"
            Set o.Icon = pic
        Case "dragicon"
            Set o.DragIcon = pic
        Case "mouseicon"
            Set o.MouseIcon = pic
        Case "disabledpicture"
            Set o.DisabledPicture = pic
        Case "downpicture"
            Set o.DownPicture = pic
        Case Else
            rs = "unsupported picture property " & propName
    End Select
    If Len(rs) = 0 Then
        If Err.Number <> 0 Then
            rs = "set on VBControl.Object raised " & CStr(Err.Number) & " " & Err.Description
            Err.Clear
        End If
    End If
    On Error GoTo 0
    TrySetPicture = rs
End Function

Private Function IsPictureProperty(ByVal propName As String) As Boolean
    Select Case LCase$(propName)
        Case "picture", "icon", "image", "dragicon", "mouseicon", "disabledpicture", "downpicture"
            IsPictureProperty = True
    End Select
End Function

' ---------------------------------------------------------------------------
' TASK-P1-040 (contract v1.4.0): PREVENT the VBControls.Add hang.
' An unregistered control part makes VBControls.Add block the whole UI thread with no
' dialog and no return (683s measured in TASK-P1-026); it is NOT interruptible in
' process, so prevention is the only cure. A part counts as registered only when the
' class@s DLL is one of the project@s .vbp Object= lines.
' ---------------------------------------------------------------------------

' Intrinsic controls are compiled into the project and therefore do NOT appear in the
' Object= lines - a naive "not in Object= -> refuse" check would wrongly reject them.
Private Function IntrinsicControlProgIds() As String
    IntrinsicControlProgIds = ",vb.checkbox,vb.combobox,vb.commandbutton,vb.data,vb.dirlistbox," & _
        "vb.drivelistbox,vb.filelistbox,vb.frame,vb.hscrollbar,vb.image,vb.label,vb.line," & _
        "vb.listbox,vb.menu,vb.ole,vb.optionbutton,vb.picturebox,vb.shape,vb.textbox,vb.timer,vb.vscrollbar,"
End Function

Private Function IsIntrinsicControl(ByVal progId As String) As Boolean
    If Len(progId) = 0 Then Exit Function
    IsIntrinsicControl = (InStr(1, IntrinsicControlProgIds(), "," & LCase$(Trim$(progId)) & ",", vbTextCompare) > 0)
End Function

Private Function RegReadStr(ByVal regPath As String) As String
    On Error Resume Next
    RegReadStr = ""
    Dim sh As Object
    Set sh = CreateObject("WScript.Shell")
    If sh Is Nothing Then Exit Function
    Dim v As Variant
    v = sh.RegRead(regPath)
    If Err.Number = 0 Then RegReadStr = CStr(v)
    Err.Clear
End Function

Private Function PathsEqual(ByVal a As String, ByVal b As String) As Boolean
    On Error Resume Next
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    If StrComp(a, b, vbTextCompare) = 0 Then PathsEqual = True
    If PathsEqual Then Exit Function
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Function
    If StrComp(fso.GetFileName(a), fso.GetFileName(b), vbTextCompare) = 0 Then PathsEqual = True
End Function

' The project@s declared ActiveX control parts = its .vbp Object= lines:
'   Object={GUID}#1.2#0; MSWINSCK.OCX
Private Function ProjectControlParts() As Collection
    Dim col As New Collection
    Set ProjectControlParts = col
    On Error Resume Next
    Dim p As String
    p = ProjectFileNameSafe()
    If Len(p) = 0 Then Exit Function
    Dim bb() As Byte, s As String
    Encoding.ReadAllBytesTo p, bb
    If Encoding.ByteCount(bb) <= 0 Then Exit Function
    Encoding.DecodeBytes bb, Encoding.DetectEncoding(bb), s
    Dim lines() As String, i As Long, ln As String
    lines = Split(Replace(s, vbCrLf, vbLf), vbLf)
    For i = LBound(lines) To UBound(lines)
        ln = Trim$(lines(i))
        If LCase$(Left$(ln, 7)) = "object=" Then
            Dim spec As String, semi As Long, guid As String, ver As String, fp As String, head As String, hash As Long
            spec = Mid$(ln, 8)
            semi = InStrRev(spec, ";")
            If semi > 0 Then
                fp = Trim$(Mid$(spec, semi + 1))
                head = Left$(spec, semi - 1)
            Else
                fp = ""
                head = spec
            End If
            hash = InStr(head, "#")
            If hash > 0 Then
                guid = Left$(head, hash - 1)
                ver = Mid$(head, hash + 1)
            Else
                guid = head
                ver = ""
            End If
            If Left$(guid, 1) = "{" And Right$(guid, 1) = "}" Then guid = Mid$(guid, 2, Len(guid) - 2)
            Dim e As Object
            Set e = CreateObject("Scripting.Dictionary")
            e("guid") = guid
            e("version") = ver
            e("path") = fp
            col.Add e
        End If
    Next
    Err.Clear
End Function

Private Function ResolvePartPath(ByVal fp As String) As String
    On Error Resume Next
    ResolvePartPath = fp
    If Len(fp) = 0 Then Exit Function
    If InStr(fp, "\") > 0 Then Exit Function
    Dim p As String, i As Long
    p = ProjectFileNameSafe()
    i = InStrRev(p, "\")
    If i > 0 Then ResolvePartPath = Left$(p, i) & fp
End Function

Private Function PartsCsv(ByVal parts As Collection) As String
    Dim i As Long, it As Object, s As String
    For i = 1 To parts.Count
        Set it = parts(i)
        If Len(s) > 0 Then s = s & ", "
        s = s & CStr(it("path"))
    Next
    PartsCsv = s
End Function

Private Function ControlPartIsRegistered(ByVal progId As String, ByRef knownCsv As String, ByRef why As String) As Boolean
    On Error Resume Next
    why = ""
    Dim parts As Collection
    Set parts = ProjectControlParts()
    knownCsv = PartsCsv(parts)
    If IsIntrinsicControl(progId) Then
        ControlPartIsRegistered = True
        Exit Function
    End If
    Dim clsid As String, dll As String
    clsid = RegReadStr("HKCR\" & progId & "\CLSID")
    If Len(clsid) = 0 Then
        why = "prog_id_not_registered"
        Exit Function
    End If
    dll = RegReadStr("HKCR\CLSID\" & clsid & "\InprocServer32")
    If Len(dll) = 0 Then dll = RegReadStr("HKCR\CLSID\" & clsid & "\LocalServer32")
    If Len(dll) = 0 Then
        why = "control_part_not_registered"
        Exit Function
    End If
    Dim i As Long, it As Object
    For i = 1 To parts.Count
        Set it = parts(i)
        If PathsEqual(ResolvePartPath(CStr(it("path"))), dll) Then
            ControlPartIsRegistered = True
            Exit Function
        End If
    Next
    why = "control_part_not_registered"
End Function

Private Function ControlPartRefusalJson(ByVal progId As String, ByVal knownCsv As String, ByVal why As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = why
    d("progId") = progId
    d("knownParts") = knownCsv
    d("advice") = "VBControls.Add hangs the whole IDE with no dialog and no timeout when the class is not a component of THIS project, and it cannot be interrupted in-process. Open Project > Components and tick the part (or pick a different control), then retry. Intrinsic VB.* controls need no registration."
    ControlPartRefusalJson = JsonStringify(d)
End Function

Private Function ToolListComponentParts() As String
    Dim parts As Collection, arr As New Collection, i As Long, it As Object, e As Object
    Set parts = ProjectControlParts()
    For i = 1 To parts.Count
        Set it = parts(i)
        Set e = CreateObject("Scripting.Dictionary")
        e("guid") = CStr(it("guid"))
        e("version") = CStr(it("version"))
        e("path") = CStr(it("path"))
        e("resolvedPath") = ResolvePartPath(CStr(it("path")))
        e("exists") = (Len(Dir$(ResolvePartPath(CStr(it("path"))))) > 0)
        arr.Add e
    Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("projectFile") = ProjectFileNameSafe()
    d("count") = parts.Count
    Set d("parts") = arr
    If parts.Count = 0 Then d("note") = "this project declares no ActiveX control parts (no .vbp Object= lines); intrinsic VB.* controls are compiled in, need no declaration and are therefore not listed"
    ToolListComponentParts = JsonStringify(d)
End Function

Private Function ToolAddControl(ByVal form As String, ByVal progId As String, ByVal wantName As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(form)
    Dim des As Object
    Set des = comp.Designer
    If des Is Nothing Then Err.Raise vbObjectError + 63, "CodeTools", "Component has no designer: " & form
    ' TASK-P1-040 (contract v1.4.0): PRE-VALIDATE. VBControls.Add blocks the whole UI
    ' thread with no dialog and no return when the class is not a component of THIS
    ' project (683s measured in TASK-P1-026) and cannot be interrupted in-process, so
    ' refusing here - before the call - is the only cure.
    Dim knownCsv As String, whyNot As String
    knownCsv = ""
    whyNot = ""
    If Not ControlPartIsRegistered(progId, knownCsv, whyNot) Then
        ToolAddControl = ControlPartRefusalJson(progId, knownCsv, whyNot)
        Exit Function
    End If
    Dim ctl As Object
    Set ctl = des.VBControls.Add(progId, Nothing, 0)
    Dim gotName As String
    gotName = ControlName(ctl)
    Dim renamed As Boolean
    renamed = False
    If Len(wantName) > 0 Then
        If StrComp(gotName, wantName, vbTextCompare) <> 0 Then
            renamed = SetControlNameTo(ctl, wantName)
            gotName = ControlName(ctl)
        End If
    End If
    Dim ctlIndex As Long
    ctlIndex = 0
    If Len(gotName) > 0 Then
        On Error Resume Next
        ctlIndex = ResolveControlIndex(des, gotName, 0)
        If Err.Number <> 0 Then
            Err.Clear
            ctlIndex = 0
        End If
        On Error GoTo 0
    End If
    If ctlIndex < 1 Then
        Dim ci As Long
        Dim ctlIt As Object
        ci = 0
        For Each ctlIt In des.VBControls
            ci = ci + 1
        Next
        ctlIndex = ci
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("form") = comp.Name
    root("progId") = progId
    root("added") = True
    root("name") = gotName
    root("index") = ctlIndex
    Err.Clear
    root("controlClass") = ctl.ClassName
    If Err.Number <> 0 Then
        Err.Clear
        root("controlClass") = ""
    End If
    If Len(wantName) > 0 Then root("nameRequested") = wantName
    If renamed Then root("renamed") = True
    ToolAddControl = JsonStringify(root)
End Function

' VBControl.Name is not readable on every VB6 build: the property read raises,
' which a naive "nm = ctl.Name" turns into an empty name and breaks every
' name-based control lookup. The designer stores the same value in the control's
' Properties collection, so fall back to that entry. Measured: without this,
' form_designer info reported name "" for controls that set_control_property
' could still address through Properties("Name").
Private Function ControlName(ByVal ctl As Object) As String
    Dim nm As String
    nm = ""
    On Error Resume Next
    nm = ctl.Name
    If Err.Number <> 0 Then
        Err.Clear
        nm = ""
    End If
    On Error GoTo 0
    If Len(nm) = 0 Then
        Dim prp As Object
        On Error Resume Next
        For Each prp In ctl.Properties
            If LCase$(prp.Name) = "name" Then
                nm = CStr(prp.Value)
                Exit For
            End If
        Next
        If Err.Number <> 0 Then nm = ""
        On Error GoTo 0
    End If
    ControlName = nm
End Function

' Writes a control's design-time Name through the Properties collection
' (VBControl.Name itself is not settable on those same builds).
Private Function SetControlNameTo(ByVal ctl As Object, ByVal nm As String) As Boolean
    Dim prp As Object
    Dim hit As Boolean
    hit = False
    On Error Resume Next
    For Each prp In ctl.Properties
        If LCase$(prp.Name) = "name" Then
            prp.Value = nm
            hit = (Err.Number = 0)
            Err.Clear
            Exit For
        End If
    Next
    On Error GoTo 0
    SetControlNameTo = hit
End Function

' Resolves a control by name (preferred) or by 1-based index.
Private Function ResolveControlIndex(ByVal des As Object, ByVal control As String, ByVal index As Long) As Long
    If Len(control) = 0 Then
        If index < 1 Then Err.Raise vbObjectError + 67, "CodeTools", "either control (name) or index (1-based) is required"
        ResolveControlIndex = index
        Exit Function
    End If
    Dim i As Long
    Dim ctl As Object
    Dim nm As String
    i = 0
    For Each ctl In des.VBControls
        i = i + 1
        nm = ControlName(ctl)
        If Len(nm) > 0 Then
            If StrComp(nm, control, vbTextCompare) = 0 Then
                ResolveControlIndex = i
                Exit Function
            End If
        End If
    Next
    Err.Raise vbObjectError + 67, "CodeTools", "Control not found by name: " & control
End Function

Private Function ToolRemoveControl(ByVal form As String, ByVal index As Long, ByVal control As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(form)
    Dim des As Object
    Set des = comp.Designer
    If des Is Nothing Then Err.Raise vbObjectError + 63, "CodeTools", "Component has no designer: " & form
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim cntBefore As Long
    Dim cntAfter As Long
    cntBefore = -1
    cntAfter = -1
    On Error Resume Next
    cntBefore = des.VBControls.Count
    If Err.Number <> 0 Then
        Err.Clear
        cntBefore = -1
    End If
    On Error GoTo 0
    Dim idx As Long
    On Error Resume Next
    Err.Clear
    idx = ResolveControlIndex(des, control, index)
    If Err.Number <> 0 Then
        Dim cfMsg As String
        cfMsg = Err.Description
        Err.Clear
        On Error GoTo 0
        root("ok") = False
        root("reason") = "control_not_found"
        root("form") = comp.Name
        root("requestedControl") = control
        root("requestedIndex") = index
        root("error") = cfMsg
        root("controlCountBefore") = cntBefore
        root("controlCountAfter") = cntBefore
        root("sideEffects") = "none: the call was refused before any control was removed"
        ToolRemoveControl = JsonStringify(root)
        Exit Function
    End If
    On Error GoTo 0
    Dim ctl As Object
    Set ctl = GetControlByIndex(des, idx)
    Dim ctlName As String
    ctlName = ControlName(ctl)
    des.VBControls.Remove ctl
    On Error Resume Next
    cntAfter = des.VBControls.Count
    If Err.Number <> 0 Then
        Err.Clear
        cntAfter = -1
    End If
    On Error GoTo 0
    ' TASK-P1-163 (R6): the receipt used to carry no 'ok' and no control counts, so the only
    ' way to confirm a deletion was a second form_designer info call. Report the verdict and
    ' the before/after counts here, and a structured 'control_not_found' failure instead of
    ' an exception when the name/index does not resolve.
    root("ok") = True
    root("removedControl") = ctlName
    root("controlCountBefore") = cntBefore
    root("controlCountAfter") = cntAfter
    If cntBefore < 0 Or cntAfter < 0 Then
        root("controlCountNote") = "VBControls.Count was not readable on this IDE; the count fields carry -1"
    End If
    ToolRemoveControl = JsonStringify(root)
End Function

Private Function ToolSetControlProperty(ByVal form As String, ByVal index As Long, ByVal control As String, ByVal propName As String, ByVal value As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(form)
    Dim des As Object
    Set des = comp.Designer
    If des Is Nothing Then Err.Raise vbObjectError + 63, "CodeTools", "Component has no designer: " & form
    Dim idx As Long
    idx = ResolveControlIndex(des, control, index)
    Dim ctl As Object
    Set ctl = GetControlByIndex(des, idx)
    Dim prop As Object
    Dim found As Boolean
    found = False
    For Each prop In ctl.Properties
        If LCase$(prop.Name) = LCase$(propName) Then
            found = True
            Exit For
        End If
    Next
    If Not found Then Err.Raise vbObjectError + 66, "CodeTools", "Control property not found: " & propName
    Dim picApplied As Boolean
    Dim picFail As String
    picApplied = False
    picFail = ""
    If IsPictureProperty(propName) Then
        ' Picture/Icon hold a StdPicture. Two routes are tried: the designer's
        ' live control instance (VBControl.Object) and a Variant put on the
        ' designer property. If neither works the call FAILS with the reason
        ' instead of storing the path as text (which raises 425 - measured).
        Dim picPath As String
        Dim pic As StdPicture
        picPath = ResolvePath(value)
        On Error Resume Next
        If Not CtlFileExists(picPath) Then
            picFail = "file not found: " & picPath
        Else
            Err.Clear
            Set pic = LoadPicture(picPath)
            If Err.Number <> 0 Then
                picFail = "LoadPicture raised " & CStr(Err.Number) & " " & Err.Description
                Err.Clear
            Else
                picFail = TrySetPicture(ctl, propName, pic)
                If Len(picFail) = 0 Then
                    picApplied = True
                Else
                    Err.Clear
                    prop.Value = pic
                    If Err.Number = 0 Then
                        picApplied = True
                        picFail = ""
                    Else
                        picFail = picFail & "; Variant put raised " & CStr(Err.Number) & " " & Err.Description
                        Err.Clear
                    End If
                End If
            End If
        End If
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
        If Not picApplied Then
            Err.Raise vbObjectError + 70, "CodeTools", "could not set " & propName & " from " & picPath & ": " & picFail
        End If
    Else
        prop.Value = ParseVariantValue(value)
    End If
    Dim ctlName As String
    ctlName = ControlName(ctl)
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("form") = comp.Name
    root("index") = idx
    root("control") = ctlName
    root("controlClass") = ctl.ClassName
    root("property") = prop.Name
    root("value") = value
    If IsPictureProperty(propName) Then root("pictureApplied") = picApplied
    ToolSetControlProperty = JsonStringify(root)
End Function

Private Function GetControlByIndex(ByVal des As Object, ByVal index As Long) As Object
    If index < 1 Then Err.Raise vbObjectError + 67, "CodeTools", "index must be >= 1"
    Dim i As Long
    i = 0
    Dim ctl As Object
    For Each ctl In des.VBControls
        i = i + 1
        If i = index Then
            Set GetControlByIndex = ctl
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 68, "CodeTools", "Control index out of range: " & CStr(index)
End Function

' =====================================================================
' property bag / add-ins / command bars / quit
' =====================================================================

Private Function ToolPropertyBag(ByVal target As String, ByVal name As String, ByVal section As String, ByVal key As String, ByVal value As String, ByVal hasValue As Boolean) As String
    ' TASK-P1-110 (D-1107-04): the schema enum is project|component, but the old code treated ANY
    ' other value as "project" and silently read/wrote the PROJECT property bag - a typo came back
    ' as an empty value with no refusal. Validate BEFORE any IDE access or write, so the refusal
    ' provably has no side effect.
    Dim tgt As String
    tgt = LCase$(Trim$(target))
    If Len(tgt) = 0 Then
        ToolPropertyBag = EnumArgRefusalJson("property_bag", "target", target, "required_argument_missing", EnumList("project", "component"))
        Exit Function
    End If
    If tgt <> "project" And tgt <> "component" Then
        ToolPropertyBag = EnumArgRefusalJson("property_bag", "target", target, "invalid_target", EnumList("project", "component"))
        Exit Function
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    If tgt = "component" Then
        Dim comp As VBIDE.VBComponent
        Set comp = FindComponent(name)
        If hasValue Then
            Dim compVal As Variant
            compVal = ParseVariantValue(value)
            comp.WriteProperty key, compVal
            root("wrote") = True
        Else
            On Error Resume Next
            root("value") = comp.ReadProperty(key)
            If Err.Number <> 0 Then root("value") = "(not set)"
            On Error GoTo 0
        End If
        root("component") = comp.Name
    Else
        Dim p As VBIDE.VBProject
        Set p = GetProject()
        If hasValue Then
            Dim pVal As Variant
            pVal = ParseVariantValue(value)
            p.WriteProperty section, key, pVal
            root("wrote") = True
        Else
            On Error Resume Next
            root("value") = p.ReadProperty(section, key)
            If Err.Number <> 0 Then root("value") = "(not set)"
            On Error GoTo 0
        End If
        root("project") = p.Name
    End If
    root("key") = key
    ' TASK-P1-114: \`section\` is honoured ONLY for target=project - VBComponent's own
    ' property-bag API (VB6EXT v5.3) is ReadProperty(Key) / WriteProperty(Key, Value) with
    ' NO section parameter at all, so a section aimed at a component can never be applied.
    ' Report it instead of dropping it silently (discipline 38); gated on a NON-EMPTY
    ' section, so the project path and the no-section path stay byte-identical.
    If tgt = "component" And Len(Trim$(section)) > 0 Then
        root("ignoredArgs") = Array("section")
        root("ignoredArgsNote") = "\`section\` addresses the PROJECT property bag only; a VBComponent's bag has no " & _
            "sections (VB6EXT v5.3: VBComponent.ReadProperty(Key) / WriteProperty(Key, Value)), so the section you passed " & _
            "was IGNORED and the value was read/written under \`key\` alone. Use target=project with the same section to " & _
            "address a sectioned project value."
    End If
    ToolPropertyBag = JsonStringify(root)
End Function

Private Function ToolListAddins() As String
    Dim arr As Object
    Set arr = New Collection
    Dim a As VBIDE.AddIn
    For Each a In m_vbe.Addins
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("progId") = a.ProgId
        d("guid") = a.Guid
        d("description") = a.Description
        d("connect") = a.Connect
        arr.Add d
    Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("count") = arr.Count
    Set root("addins") = arr
    ToolListAddins = JsonStringify(root)
End Function

Private Function ToolSetAddinState(ByVal progId As String, ByVal connect As Boolean) As String
    Dim a As VBIDE.AddIn
    Dim found As Boolean
    found = False
    For Each a In m_vbe.Addins
        If LCase$(a.ProgId) = LCase$(progId) Then
            found = True
            Exit For
        End If
    Next
    If Not found Then Err.Raise vbObjectError + 69, "CodeTools", "Add-in not found: " & progId
    a.Connect = connect
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("progId") = a.ProgId
    root("connect") = a.Connect
    ToolSetAddinState = JsonStringify(root)
End Function

' Card TASK-P1-006 (i18n prerequisite): expose the language-INDEPENDENT identity of
' a menu control. Every property read is guarded because several of them raise for
' particular control types (separators, custom bars), and defaults are written first
' so the JSON shape stays stable even when a read is skipped.
Private Sub AddCtrlMeta(ByVal c As Object, ByVal e As Object, ByVal pfx As String)
    e("caption") = ""
    e("enabled") = False
    e("visible") = False
    e("id") = 0
    e("tag") = ""
    e("shortcutText") = ""
    e("type") = 0
    e("beginGroup") = False
    e("path") = pfx
    On Error Resume Next
    e("caption") = c.Caption
    e("enabled") = c.Enabled
    e("visible") = c.Visible
    e("id") = c.Id
    e("tag") = c.Tag
    e("shortcutText") = c.ShortcutText
    e("type") = c.Type
    e("beginGroup") = c.BeginGroup
    e("path") = pfx & CStr(e("caption"))
    On Error GoTo 0
End Sub

Private Function ToolCommandBars(ByVal action As String, ByVal bar As String, ByVal control As String) As String
    ' TASK-P1-111 (S4): EVERY action other than "execute" silently fell through to the bar LIST, so a
    ' typo was answered with the 34 menu bars as if the call had succeeded (the same silent-acceptance
    ' shape as TASK-P1-110's property_bag). Validate the declared enum FIRST, so the refusal is
    ' structured, names allowedValues and is provably free of IDE access.
    Dim cbAct As String
    cbAct = LCase$(Trim$(action))
    If Len(cbAct) = 0 Then
        ToolCommandBars = EnumArgRefusalJson("command_bars", "action", action, "required_argument_missing", EnumList("list", "execute"))
        Exit Function
    End If
    If cbAct <> "list" And cbAct <> "execute" Then
        ToolCommandBars = EnumArgRefusalJson("command_bars", "action", action, "invalid_action", EnumList("list", "execute"))
        Exit Function
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    ' TASK-P1-069: address a menu item by Id through the EXISTING control
    ' parameter, so this tool's inputSchema does not change: "#746", "id:746"
    ' or a bare number. VB6 Ids are NOT unique (TASK-P1-006 measured id=746
    ' alone covering 986 nodes), so the id path executes only when EXACTLY ONE
    ' ENABLED control matches and otherwise reports the ambiguity together with
    ' its candidates instead of guessing.
    Dim ctlSpec As String, idText As String, ctlId As Long, byId As Boolean, ctlCap As String
    ctlSpec = Trim$(control)
    ctlId = 0
    byId = False
    ctlCap = ""
    If Len(ctlSpec) > 0 Then
        idText = ctlSpec
        If Left$(idText, 1) = "#" Then
            byId = True
            idText = Trim$(Mid$(idText, 2))
        ElseIf LCase$(Left$(idText, 3)) = "id:" Then
            byId = True
            idText = Trim$(Mid$(idText, 4))
        ElseIf IsNumeric(Left$(idText, 1)) Then
            byId = True
        End If
        If byId Then
            ' "#746" = id only. "#746 部件" / "#746:部件" = id AND caption. The
            ' caption conjunct matters: TASK-P1-069 measured id=746 with 525
            ' ENABLED matches across the menus (6 inside the 工程 submenu alone),
            ' so an id on its own cannot name a control - the project's own
            ' REQ-20260921-016 rule is that an Id is never used on its own.
            Dim cut As Long, sp As Long, cp As Long
            sp = InStr(idText, " ")
            cp = InStr(idText, ":")
            cut = 0
            If sp > 0 Then cut = sp
            If cp > 0 Then
                If cut = 0 Or cp < cut Then cut = cp
            End If
            If cut > 0 Then
                ctlCap = Trim$(Mid$(idText, cut + 1))
                idText = Trim$(Left$(idText, cut - 1))
            End If
            If Not IsNumeric(idText) Then
                Err.Raise vbObjectError + 72, "CodeTools", "control looks like an id but its number is missing: " & control
            End If
            ctlId = CLng(Val(idText))
        End If
    End If
    If cbAct = "execute" Then
        If Len(ctlSpec) = 0 Then Err.Raise vbObjectError + 70, "CodeTools", "control is required for action=execute (a caption, or an id as #746 / id:746)"
        Dim ctrl2 As Object
        Dim idHits As Long, idCands As String
        idHits = 0
        idCands = ""
        If byId Then
            Set ctrl2 = FindBarControlByIdOnly(m_vbe, bar, ctlCap, ctlId, idHits, idCands)
            If ctrl2 Is Nothing Then
                If idHits = 0 Then
                    Err.Raise vbObjectError + 73, "CodeTools", "no ENABLED menu control with id=" & CStr(ctlId) & IIf(Len(ctlCap) > 0, " and caption like """ & ctlCap & """", "") & " (bar filter=""" & bar & """)"
                Else
                    Err.Raise vbObjectError + 74, "CodeTools", "id=" & CStr(ctlId) & " is AMBIGUOUS: " & CStr(idHits) & " enabled controls match; narrow it with bar= (a bar name OR a menu path segment) and/or the caption as ""#746 <caption>"". Candidates: " & idCands
                End If
            End If
        Else
            Set ctrl2 = FindBarControl(m_vbe, bar, control)
        End If
        If ctrl2 Is Nothing Then
            ' fallback: substring prefix search across every bar and popup
            Dim bars2 As Object
            Set bars2 = m_vbe.CommandBars
            Dim b2 As Object
            Dim found2 As Boolean
            found2 = False
            Dim nb2 As Long
            nb2 = bars2.Count
            Dim bi2 As Long
            For bi2 = 1 To nb2
                Set b2 = bars2.Item(bi2)
                If Not b2 Is Nothing Then
                    If Len(bar) = 0 Or MatchCand(b2.Name, bar) Then
                        Dim nc2 As Long
                        nc2 = b2.Controls.Count
                        Dim ci2 As Long
                        Dim c2 As Object
                        For ci2 = 1 To nc2
                            Set c2 = b2.Controls.Item(ci2)
                            If Not c2 Is Nothing Then
                                If Len(c2.Caption) > 0 Then
                                    If InStr(1, c2.Caption, control, vbTextCompare) = 1 Then
                                        If c2.Enabled Then
                                            c2.Execute
                                            found2 = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                End If
                If found2 Then Exit For
            Next
            If Not found2 Then Err.Raise vbObjectError + 71, "CodeTools", "Menu control not found: " & control
        Else
            Dim ok2 As Boolean
            ok2 = True
            On Error Resume Next
            ok2 = ctrl2.Enabled
            On Error GoTo 0
            If Not ok2 Then Err.Raise vbObjectError + 71, "CodeTools", "Menu control is disabled in the current IDE state: " & control
            ctrl2.Execute
        End If
        root("executed") = control
        root("byId") = byId
        root("id") = ctlId
        If Len(idCands) > 0 Then root("idCandidates") = idCands
    Else
        Dim arr As Object
        Set arr = New Collection
        Dim bars As Object
        Set bars = m_vbe.CommandBars
        Dim b As Object
        Dim nbL As Long
        nbL = bars.Count
        Dim biL As Long
        For biL = 1 To nbL
            Set b = bars.Item(biL)
            If Not b Is Nothing Then
                If Len(bar) = 0 Or MatchCand(b.Name, bar) Then
                    Dim d As Object
                    Set d = CreateObject("Scripting.Dictionary")
                    ' Card TASK-P1-006: the bar NAME is localized too, so carry id/type
                    ' alongside it. Defaults first, then guarded reads.
                    d("bar") = b.Name
                    d("id") = 0
                    d("type") = 0
                    d("visible") = False
                    d("builtIn") = False
                    d("position") = 0
                    On Error Resume Next
                    d("id") = b.Id
                    d("type") = b.Type
                    d("visible") = b.Visible
                    d("builtIn") = b.BuiltIn
                    d("position") = b.Position
                    On Error GoTo 0
                    Dim items As Object
                    Set items = New Collection
                    Dim c As Object
                    Dim ncL As Long
                    ncL = b.Controls.Count
                    Dim ciL As Long
                    For ciL = 1 To ncL
                        On Error Resume Next
                        Set c = b.Controls.Item(ciL)
                        If Not c Is Nothing Then
                            Dim e As Object
                            Set e = CreateObject("Scripting.Dictionary")
                            AddCtrlMeta c, e, b.Name & ">"
                            Dim subs As Object
                            Set subs = New Collection
                            Dim sc As Object
                            Dim scCol As Object
                            Set scCol = c.Controls
                            If Not scCol Is Nothing Then
                                Dim nsL As Long
                                nsL = scCol.Count
                                Dim siL As Long
                                For siL = 1 To nsL
                                    Set sc = scCol.Item(siL)
                                    If Not sc Is Nothing Then
                                        Dim se As Object
                                        Set se = CreateObject("Scripting.Dictionary")
                                        AddCtrlMeta sc, se, b.Name & ">" & CStr(e("caption")) & ">"
                                        subs.Add se
                                    End If
                                Next
                            End If
                            If subs.Count > 0 Then Set e("subControls") = subs
                            items.Add e
                        End If
                        On Error GoTo 0
                    Next
                    d("controlCount") = items.Count
                    Set d("controls") = items
                    arr.Add d
                End If
            End If
        Next
        root("barCount") = arr.Count
        Set root("bars") = arr
    End If
    ToolCommandBars = JsonStringify(root)
End Function

Private Function ToolQuitIde() As String
    If Not m_server Is Nothing Then m_server.RequestQuit
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("quitting") = True
    ToolQuitIde = JsonStringify(root)
End Function


Public Sub SetServer(ByVal server As Object)
    Set m_server = server
End Sub

' =====================================================================
' VBIDE helpers
' =====================================================================

Private Function GetProject() As VBIDE.VBProject
    If m_vbe Is Nothing Then Err.Raise vbObjectError + 1, "CodeTools", "IDE not connected"
    ' TASK-P1-091: "no project is open" is a LEGAL IDE state, not a fault. Tag it with
    ' MCPREASON: so CallTool's Fail handler answers with the structured refusal
    ' (ok:false + reason:"no_project") for EVERY caller at once, instead of wrapping the
    ' raw VB error into ok:false/unexpected_error. Nothing else changes: with a project
    ' open this raise is unreachable, so those receipts stay byte-identical.
    If m_vbe.VBProjects.Count = 0 Then Err.Raise vbObjectError + 2, "CodeTools", "MCPREASON:no_project|open or create a project first (File > Open Project)|the VB6 IDE has no open project, so this tool has nothing to operate on"
    Set GetProject = m_vbe.ActiveVBProject
End Function

Private Function FindComponent(ByVal name As String) As VBIDE.VBComponent
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        If LCase$(comp.Name) = LCase$(name) Then
            Set FindComponent = comp
            Exit Function
        End If
    Next
    ' TASK-P1-101: "the component you named is not in this project" is CALLER INPUT, not an
    ' add-in fault. It used to be a bare raise (+20) that CallToolInner could only wrap as
    ' ok:false / reason:"unexpected_error" / vbErrorCode:-2147221484 for EVERY one of the 36
    ' call sites (41 tools once the helpers are counted). Capture the diagnostics the caller
    ' needs, then tag the raise so the MCPREASON channel answers with a branchable refusal.
    ' This is the ONLY place affected: the success path (component found) does not change at
    ' all, and no other refusal is touched.
    CnfCapture name, p
    Err.Raise vbObjectError + 20, "CodeTools", "MCPREASON:component_not_found"
End Function

' TASK-P1-101: capture what the caller needs to fix the name. Lives in module state because a
' component name may contain anything; every read is defensive (a diagnostic must never turn a
' refusal into an unexpected error - TASK-P1-079).
Private Sub CnfCapture(ByVal name As String, ByVal p As VBIDE.VBProject)
    Set m_cnfNames = New Collection
    Set m_cnfNear = New Collection
    m_cnfTried = name
    m_cnfCount = 0
    m_cnfTrimmed = ""
    m_cnfCollapsed = ""
    On Error Resume Next
    Dim lc As String, ltrim As String, lcol As String
    lc = LCase$(name)
    ltrim = LCase$(Trim$(name))
    lcol = LCase$(CollapseWs(name))
    Dim comp As VBIDE.VBComponent
    Dim n As String, nl As String, nd As Object
    For Each comp In p.VBComponents
        n = comp.Name
        nl = LCase$(n)
        m_cnfCount = m_cnfCount + 1
        If m_cnfCount <= 100 Then m_cnfNames.Add n
        If nl = ltrim Then
            If Len(m_cnfTrimmed) = 0 Then m_cnfTrimmed = n
        End If
        If nl = lcol Then
            If Len(m_cnfCollapsed) = 0 Then m_cnfCollapsed = n
        End If
        If Len(lc) >= 2 Then
            If InStr(1, nl, lc, vbBinaryCompare) > 0 Then CnfNearAdd n, "an existing component's name CONTAINS the name you passed - you may have passed a prefix or a fragment"
        End If
        If Len(nl) >= 2 Then
            If InStr(1, lc, nl, vbBinaryCompare) > 0 Then CnfNearAdd n, "the name you passed CONTAINS an existing component's name - you may have passed a longer decorated name"
        End If
    Next
    If Len(m_cnfTrimmed) > 0 Then CnfNearAdd m_cnfTrimmed, "matches after Trim$: the name differs only by SURROUNDING whitespace"
    If Len(m_cnfCollapsed) > 0 Then CnfNearAdd m_cnfCollapsed, "matches after collapsing INNER whitespace (and trimming)"
End Sub

Private Sub CnfNearAdd(ByVal n As String, ByVal why As String)
    If Len(Trim$(n)) = 0 Then Exit Sub
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("name") = n
    d("why") = why
    m_cnfNear.Add d
End Sub

Private Function CollapseWs(ByVal s As String) As String
    Dim t As String, i As Long, ch As String, out As String, sp As Boolean
    t = Trim$(s)
    sp = False
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If ch = " " Or ch = vbTab Then
            If Not sp Then out = out & " "
            sp = True
        Else
            out = out & ch
            sp = False
        End If
    Next
    CollapseWs = out
End Function

Private Function GetCodePane(ByVal component As String) As VBIDE.CodePane
    If Len(component) = 0 Then
        Set GetCodePane = m_vbe.ActiveCodePane
        If GetCodePane Is Nothing Then Err.Raise vbObjectError + 72, "CodeTools", "No active code pane"
    Else
        Dim comp As VBIDE.VBComponent
        Set comp = FindComponent(component)
        Dim cm As VBIDE.CodeModule
        Set cm = comp.CodeModule
        If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & component
        Set GetCodePane = cm.CodePane
    End If
End Function

' Card TASK-P1-007 (i18n): capability -> language-independent Id + shortcut.
' Ids/shortcuts measured on the Chinese VB6 IDE (内部命令栏清单).
' Returning 0 means "no Id may be used": 'make' shares the ambiguous id=746 with
' start_full_compile and has NO shortcut, so it must stay caption-based.
Private Function CmdId(ByVal key As String, ByRef shortcutOut As String) As Long
    shortcutOut = ""
    Select Case LCase$(key)
        Case "run", "start", "continue", "resume"
            CmdId = 186: shortcutOut = "F5"
        Case "break"
            CmdId = 189: shortcutOut = "Ctrl+Break"
        Case "end", "stop"
            CmdId = 228
        Case "step_into", "stepinto"
            CmdId = 188: shortcutOut = "F8"
        Case "step_over", "stepover"
            CmdId = 194: shortcutOut = "Shift+F8"
        Case "step_out", "stepout"
            CmdId = 2559: shortcutOut = "Ctrl+Shift+F8"
        Case "run_to_cursor", "runtocursor"
            CmdId = 1811: shortcutOut = "Ctrl+F8"
        Case "start_full_compile", "full_compile", "fullcompile"
            ' id=746 is ambiguous across 986 nodes, so the shortcut is mandatory
            CmdId = 746: shortcutOut = "Ctrl+F5"
        Case "toggle", "toggle_breakpoint"
            CmdId = 51: shortcutOut = "F9"
        Case "clear_all", "clear_all_breakpoints"
            CmdId = 579: shortcutOut = "Ctrl+Shift+F9"
        Case "addins_menu"
            CmdId = 30038
        Case Else
            ' make / command_bars execute / unknown: no safe Id
            CmdId = 0
    End Select
End Function

Private Function FindCmd(ByVal key As String, ByVal barCands As String, ByVal ctrlCands As String) As Object
    Dim sc As String
    Dim i As Long
    sc = ""
    i = CmdId(key, sc)
    Set FindCmd = FindBarControlById(m_vbe, barCands, ctrlCands, i, sc)
End Function

Private Function GetIdeState() As String
    On Error Resume Next
    Dim startCtrl As Object
    Dim breakCtrl As Object
    Dim endCtrl As Object
    Set startCtrl = FindCmd("start", "Debug|" & ChrW$(&H8C03) & ChrW$(&H8BD5), "Start|" & ChrW$(&H542F) & ChrW$(&H52A8) & "|Continue|" & ChrW$(&H7EE7) & ChrW$(&H7EED))
    If startCtrl Is Nothing Then Set startCtrl = m_cachedStart Else Set m_cachedStart = startCtrl
    Set breakCtrl = FindCmd("break", "Debug|" & ChrW$(&H8C03) & ChrW$(&H8BD5), "Break|" & ChrW$(&H4E2D) & ChrW$(&H65AD))
    If breakCtrl Is Nothing Then Set breakCtrl = m_cachedBreak Else Set m_cachedBreak = breakCtrl
    Set endCtrl = FindCmd("end", "Debug|" & ChrW$(&H8C03) & ChrW$(&H8BD5), "End|" & ChrW$(&H7ED3) & ChrW$(&H675F))
    If endCtrl Is Nothing Then Set endCtrl = m_cachedEnd Else Set m_cachedEnd = endCtrl
    Dim startEnabled As Boolean
    Dim breakEnabled As Boolean
    Dim endEnabled As Boolean
    startEnabled = False
    breakEnabled = False
    endEnabled = False
    If Not startCtrl Is Nothing Then startEnabled = startCtrl.Enabled
    If Not breakCtrl Is Nothing Then breakEnabled = breakCtrl.Enabled
    If Not endCtrl Is Nothing Then endEnabled = endCtrl.Enabled
    ' order matters: in break mode the Start button becomes Continue
    ' (继续) and is enabled, so check break first, then end, then start
    ' Card TASK-P1-007 (i18n): EbMode is a VBA6 debug-engine signal that does not
    ' depend on any localized text: 2 = break context, 0 = no debug context.
    If IdeEbMode() = 2 Then
        GetIdeState = "break"
        Exit Function
    End If
    If breakEnabled Then
        GetIdeState = "run"
    ElseIf endEnabled Then
        GetIdeState = "break"
    ElseIf startEnabled Then
        GetIdeState = "design"
    Else
        ' run mode suspends the CommandBars object model: fall back to
        ' the IDE main window caption ([运行] / [设计] / [中断])
        ' Card TASK-P1-007 (i18n): no localized caption sniffing any more. When
        ' every CommandBars lookup fails the object model is suspended, which
        ' happens ONLY while the program runs.
        GetIdeState = "run"
    End If
End Function

Private Function ProjectDir() As String
    On Error GoTo np
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If Len(p.FileName) = 0 Then Exit Function
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ProjectDir = fso.GetParentFolderName(p.FileName)
    Exit Function
np:
    ProjectDir = ""
End Function

Private Function ComponentListArr() As Object
    Dim c As Object
    Set c = New Collection
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = comp.Name
        d("type") = ComponentTypeName(comp.Type)
        On Error Resume Next
        Err.Clear
        d("file") = comp.FileNames(1)
        On Error GoTo 0
        Dim cm As VBIDE.CodeModule
        Set cm = comp.CodeModule
        If Not cm Is Nothing Then
            d("lines") = cm.CountOfLines
            d("declarationLines") = cm.CountOfDeclarationLines
        End If
        c.Add d
    Next
    Set ComponentListArr = c
End Function

Private Function ReferenceListArr() As Object
    Dim c As Object
    Set c = New Collection
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim ref As VBIDE.Reference
    For Each ref In p.References
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        On Error Resume Next
        Err.Clear
        d("name") = ref.Name
        If Err.Number <> 0 Then
            d("broken") = True
        Else
            d("description") = ref.Description
            d("guid") = ref.Guid
            d("major") = ref.Major
            d("minor") = ref.Minor
            d("fullPath") = ref.FullPath
            d("builtIn") = ref.BuiltIn
            d("type") = RefKindName(ref.Type)
        End If
        On Error GoTo 0
        c.Add d
    Next
    Set ReferenceListArr = c
End Function

Private Function ProcsOf(ByVal comp As VBIDE.VBComponent) As Object
    Dim c As Object
    Set c = New Collection
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then
        Set ProcsOf = c
        Exit Function
    End If
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim line As Long
    For line = 1 To cm.CountOfLines
        MaybePump
        Dim k As Integer
        For k = 0 To 3
            Dim kind As Long
            Dim kindName As String
            Select Case k
                Case 0
                    kind = vbext_pk_Proc
                    kindName = "procedure"
                Case 1
                    kind = vbext_pk_Get
                    kindName = "property-get"
                Case 2
                    kind = vbext_pk_Let
                    kindName = "property-let"
                Case 3
                    kind = vbext_pk_Set
                    kindName = "property-set"
            End Select
            Dim pname As String
            pname = cm.ProcOfLine(line, kind)
            If Len(pname) > 0 Then
                If Not seen.Exists(pname) Then
                    seen(pname) = True
                    Dim d As Object
                    Set d = CreateObject("Scripting.Dictionary")
                    d("name") = pname
                    d("kind") = kindName
                    d("startLine") = cm.ProcStartLine(pname, kind)
                    d("lineCount") = cm.ProcCountLines(pname, kind)
                    d("bodyLine") = cm.ProcBodyLine(pname, kind)
                    c.Add d
                End If
            End If
        Next k
    Next line
    Set ProcsOf = c
End Function

Private Function ProjectTypeName(ByVal p As VBIDE.VBProject) As String
    On Error Resume Next
    Dim t As vbext_ProjectType
    t = p.Type
    If Err.Number <> 0 Then
        ProjectTypeName = "unknown"
        Exit Function
    End If
    On Error GoTo 0
    Select Case t
        Case vbext_pt_StandardExe
            ProjectTypeName = "StandardExe"
        Case vbext_pt_ActiveXExe
            ProjectTypeName = "ActiveXExe"
        Case vbext_pt_ActiveXDll
            ProjectTypeName = "ActiveXDll"
        Case vbext_pt_ActiveXControl
            ProjectTypeName = "ActiveXControl"
        Case Else
            ProjectTypeName = "Unknown(" & CStr(t) & ")"
    End Select
End Function

Private Function ComponentTypeName(ByVal t As vbext_ComponentType) As String
    Select Case t
        Case vbext_ct_StdModule
            ComponentTypeName = "Module"
        Case vbext_ct_ClassModule
            ComponentTypeName = "Class"
        Case vbext_ct_MSForm
            ComponentTypeName = "Form"
        Case vbext_ct_ResFile
            ComponentTypeName = "Resource"
        Case vbext_ct_VBForm
            ComponentTypeName = "VBForm"
        Case vbext_ct_VBMDIForm
            ComponentTypeName = "MDIForm"
        Case vbext_ct_PropPage
            ComponentTypeName = "PropertyPage"
        Case vbext_ct_UserControl
            ComponentTypeName = "UserControl"
        Case vbext_ct_DocObject
            ComponentTypeName = "UserDocument"
        Case vbext_ct_RelatedDocument
            ComponentTypeName = "RelatedDocument"
        Case vbext_ct_ActiveXDesigner
            ComponentTypeName = "Designer"
        Case Else
            ComponentTypeName = "Unknown(" & CStr(t) & ")"
    End Select
End Function

Private Function StartModeName(ByVal p As VBIDE.VBProject) As String
    On Error Resume Next
    Dim t As vbext_ProjectStartMode
    t = p.StartMode
    If Err.Number <> 0 Then
        StartModeName = "unknown"
        Exit Function
    End If
    On Error GoTo 0
    If t = vbext_psm_OleServer Then
        StartModeName = "OleServer"
    Else
        StartModeName = "StandAlone"
    End If
End Function

Private Function DisplayModelName() As String
    On Error Resume Next
    If m_vbe.DisplayModel = vbext_dm_SDI Then
        DisplayModelName = "SDI"
    Else
        DisplayModelName = "MDI"
    End If
End Function

Private Function CodePaneViewName(ByVal v As Long) As String
    If v = vbext_cv_ProcedureView Then
        CodePaneViewName = "procedure"
    Else
        CodePaneViewName = "full_module"
    End If
End Function

Private Function RefKindName(ByVal t As Long) As String
    If t = vbext_rk_Project Then
        RefKindName = "Project"
    Else
        RefKindName = "TypeLib"
    End If
End Function

Private Function MemberTypeName(ByVal t As Long) As String
    Select Case t
        Case vbext_mt_Method
            MemberTypeName = "Method"
        Case vbext_mt_Property
            MemberTypeName = "Property"
        Case vbext_mt_Variable
            MemberTypeName = "Variable"
        Case vbext_mt_Event
            MemberTypeName = "Event"
        Case vbext_mt_Const
            MemberTypeName = "Const"
        Case Else
            MemberTypeName = "Unknown"
    End Select
End Function

Private Function ScopeName(ByVal s As Long) As String
    Select Case s
        Case vbext_Private
            ScopeName = "Private"
        Case vbext_Public
            ScopeName = "Public"
        Case vbext_Friend
            ScopeName = "Friend"
        Case Else
            ScopeName = "Unknown"
    End Select
End Function

Private Function WindowTypeName(ByVal t As vbext_WindowType) As String
    Select Case t
        Case vbext_wt_CodeWindow
            WindowTypeName = "CodeWindow"
        Case vbext_wt_Designer
            WindowTypeName = "Designer"
        Case vbext_wt_Browser
            WindowTypeName = "Browser"
        Case vbext_wt_Watch
            WindowTypeName = "Watch"
        Case vbext_wt_Locals
            WindowTypeName = "Locals"
        Case vbext_wt_Immediate
            WindowTypeName = "Immediate"
        Case vbext_wt_ProjectWindow
            WindowTypeName = "ProjectWindow"
        Case vbext_wt_PropertyWindow
            WindowTypeName = "PropertyWindow"
        Case vbext_wt_Find
            WindowTypeName = "Find"
        Case vbext_wt_FindReplace
            WindowTypeName = "FindReplace"
        Case vbext_wt_Toolbox
            WindowTypeName = "Toolbox"
        Case vbext_wt_LinkedWindowFrame
            WindowTypeName = "LinkedWindowFrame"
        Case vbext_wt_MainWindow
            WindowTypeName = "MainWindow"
        Case vbext_wt_Preview
            WindowTypeName = "Preview"
        Case vbext_wt_ColorPalette
            WindowTypeName = "ColorPalette"
        Case vbext_wt_ToolWindow
            WindowTypeName = "ToolWindow"
        Case Else
            WindowTypeName = "Unknown"
    End Select
End Function

Private Function WindowStateName(ByVal s As vbext_WindowState) As String
    Select Case s
        Case vbext_ws_Minimize
            WindowStateName = "Minimized"
        Case vbext_ws_Maximize
            WindowStateName = "Maximized"
        Case Else
            WindowStateName = "Normal"
    End Select
End Function

' =====================================================================
' file system helpers
' =====================================================================

' Card TASK-P1-015: two paths name the same file when their CANONICAL identities
' match - junction/symlink alias, relative form, separator style and case all
' collapse to one identity. Never compare spellings (LCase$ equality is what let
' a project open in the IDE be handed to an isolated /make that cannot take it).
Private Function PathsSame(ByVal a As String, ByVal b As String) As Boolean
    PathsSame = False
    If Len(a) = 0 Or Len(b) = 0 Then Exit Function
    Dim ca As String, cb As String
    ca = Build.CanonicalPath(a)
    cb = Build.CanonicalPath(b)
    If Len(ca) = 0 Or Len(cb) = 0 Then Exit Function
    PathsSame = (StrComp(ca, cb, vbTextCompare) = 0)
End Function

Private Function ResolvePath(ByVal p As String) As String
    ' Card TASK-P1-014: decide "is this absolute?" from the SHAPE of the input, NEVER by
    ' comparing it with FileSystemObject's normalisation. The old code did
    '     absP = fso.GetAbsolutePathName(p) : If absP = p Then absolute
    ' so ANY normalisation difference made an absolute path look relative. The
    ' decisive case is a DOUBLED separator ("D:\\a\\b.vbp", what an over-escaped
    ' caller sends): GetAbsolutePathName collapses it, the comparison failed, the path
    ' was concatenated onto the active project directory, and the resulting
    ' non-existent path was handed to VB6.EXE /make - which HUNG and saturated the CPU
    ' until the caller was killed by hand.
    If Len(p) = 0 Then Err.Raise vbObjectError + 30, "CodeTools", "path is required"
    ' Extended-length / device prefixes are returned verbatim: those paths must not be
    ' normalised at all (collapsing their separators changes their meaning).
    If Left$(p, 4) = "\\?\" Or Left$(p, 4) = "\\.\" Then
        ResolvePath = p
        Exit Function
    End If
    ' unify separators and collapse runs of them. A LEADING "\\" is the UNC marker
    ' and must survive collapsing - folding it away made a UNC path look relative,
    ' which is the same class of bug this card fixes.
    Dim i As Long, ch As String, t As String, prevSep As Boolean
    Dim startI As Long
    startI = 1
    t = ""
    prevSep = False
    If Left$(p, 2) = "\\" Then
        t = "\\"
        prevSep = True
        startI = 3
    End If
    For i = startI To Len(p)
        ch = Mid$(p, i, 1)
        If ch = "/" Then ch = "\"
        If ch = "\" Then
            If Not prevSep Then t = t & "\"
            prevSep = True
        Else
            t = t & ch
            prevSep = False
        End If
    Next
    ' Windows cannot address a final component ending in a space or a dot: reject it
    ' explicitly instead of letting the ambiguity reach the compiler.
    Dim lastComp As String, lastCh As String
    lastComp = Mid$(t, InStrRev(t, "\") + 1)
    If Len(lastComp) > 0 Then
        lastCh = Right$(lastComp, 1)
        If (lastCh = " " Or lastCh = ".") And lastComp <> "." And lastComp <> ".." Then
            Err.Raise vbObjectError + 30, "CodeTools", "invalid path (final component ends with a space or dot): " & p
        End If
    End If
    ' absolute by shape: drive prefix "X:" or UNC leading "\\"
    Dim isAbs As Boolean, driveRel As Boolean
    isAbs = False
    driveRel = False
    If Len(t) >= 2 Then
        If Mid$(t, 2, 1) = ":" Then
            isAbs = True
            ' "C:name" (no separator after the colon) is relative to the CURRENT
            ' directory on drive C: - a documented Windows form. Delegate exactly that
            ' case to FSO, which implements the semantic.
            If Len(t) = 2 Then
                driveRel = True
            ElseIf Mid$(t, 3, 1) <> "\" Then
                driveRel = True
            End If
        End If
    End If
    If Left$(t, 2) = "\\" Then isAbs = True
    If isAbs Then
        If driveRel Then
            Dim fso2 As Object
            Set fso2 = CreateObject("Scripting.FileSystemObject")
            ResolvePath = fso2.GetAbsolutePathName(t)
        Else
            ResolvePath = t
        End If
        Exit Function
    End If
    ResolvePath = CombinePath(ProjectDir(), t)
End Function

Private Function CombinePath(ByVal a As String, ByVal b As String) As String
    If a = "" Then
        CombinePath = b
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    CombinePath = fso.BuildPath(a, b)
End Function

Private Sub EnsureDir(ByVal dirPath As String)
    If Len(dirPath) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(dirPath) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(dirPath)
    If Len(parent) > 0 And parent <> dirPath Then
        If Not fso.FolderExists(parent) Then EnsureDir parent
    End If
    fso.CreateFolder dirPath
End Sub

Private Function ReadAllBytes(ByVal path As String) As Byte()
    Dim fnum As Integer
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    Dim b() As Byte
    If LOF(fnum) > 0 Then
        ReDim b(0 To LOF(fnum) - 1)
        Get #fnum, , b
        ReadAllBytes = b
    Else
        ReadAllBytes = b
    End If
    Close #fnum
End Function

Private Function ReadTextAuto(ByVal path As String) As String
    Dim b() As Byte
    b = ReadAllBytes(path)
    Dim n As Long
    n = UBound(b) - LBound(b) + 1
    If n <= 0 Then Exit Function
    If n >= 2 Then
        If b(LBound(b)) = &HFF And b(LBound(b) + 1) = &HFE Then
            ReadTextAuto = Utf16DecodeBytes(b)
            Exit Function
        End If
    End If
    Dim start As Long
    start = 0
    If n >= 3 Then
        If b(LBound(b)) = &HEF And b(LBound(b) + 1) = &HBB And b(LBound(b) + 2) = &HBF Then start = 3
    End If
    Dim subArr() As Byte
    If start > 0 Then
        ReDim subArr(0 To n - start - 1)
        Dim k As Long
        For k = 0 To n - start - 1
            MaybePump
            subArr(k) = b(LBound(b) + start + k)
        Next
    Else
        subArr = b
    End If
    Dim txt As String
    If TryUtf8Decode(subArr, txt) Then
        ReadTextAuto = txt
    Else
        ReadTextAuto = AnsiDecodeBytes(subArr)
    End If
End Function

Private Function IsVbFileExt(ByVal path As String) As Boolean
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim ext As String
    ext = LCase$(fso.GetExtensionName(path))
    Select Case ext
        Case "bas", "cls", "frm", "vbp", "vbg", "ctl", "pag", "dsr", "dob", "dox", "res", "inc"
            IsVbFileExt = True
    End Select
End Function

Private Function IsProjectFile(ByVal full As String) As Boolean
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If p Is Nothing Then Exit Function
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        Dim f As String
        f = comp.FileNames(1)
        If LCase$(f) = LCase$(full) Then
            IsProjectFile = True
            Exit Function
        End If
    Next
End Function

Private Function RelPath(ByVal base As String, ByVal full As String) As String
    If Len(base) = 0 Then
        RelPath = full
        Exit Function
    End If
    Dim b As String
    b = base
    If Right$(b, 1) <> "\" Then b = b & "\"
    If LCase$(Left$(full, Len(b))) = LCase$(b) Then
        RelPath = Mid$(full, Len(b) + 1)
    Else
        RelPath = full
    End If
End Function

Private Function MatchPattern(ByVal name As String, ByVal patterns As Variant) As Boolean
    Dim i As Long
    For i = LBound(patterns) To UBound(patterns)
        Dim p As String
        p = Trim$(patterns(i))
        If Len(p) = 0 Then p = "*"
        If p = "*" Or LCase$(name) Like LCase$(p) Then
            MatchPattern = True
            Exit Function
        End If
    Next
End Function

Private Sub WalkDir(ByVal fso As Object, ByVal folder As Object, ByVal patterns As Variant, ByVal results As Object, ByVal base As String, ByVal recursive As Boolean, ByVal maxEntries As Long, ByRef count As Long, ByRef truncated As Boolean)
    If truncated Then Exit Sub
    Dim f As Object
    For Each f In folder.Files
        If count >= maxEntries Then
            truncated = True
            Exit Sub
        End If
        If MatchPattern(f.Name, patterns) Then
            results.Add RelPath(base, f.Path)
            count = count + 1
        End If
    Next
    If Not recursive Then Exit Sub
    Dim subf As Object
    For Each subf In folder.SubFolders
        If truncated Then Exit Sub
        Dim nm As String
        nm = LCase$(subf.Name)
        If nm <> ".git" And nm <> ".svn" And nm <> ".hg" And nm <> "node_modules" Then
            WalkDir fso, subf, patterns, results, base, recursive, maxEntries, count, truncated
        End If
    Next
End Sub

Private Sub SearchWalk(ByVal fso As Object, ByVal folder As Object, ByVal patterns As Variant, ByVal re As Object, ByVal query As String, ByVal isRegex As Boolean, ByVal caseSensitive As Boolean, ByVal maxResults As Long, ByVal results As Object, ByRef count As Long, ByRef truncated As Boolean, ByRef filesScanned As Long, ByVal base As String)
    If truncated Then Exit Sub
    Dim f As Object
    For Each f In folder.Files
        If truncated Then Exit Sub
        If filesScanned >= 2000 Then
            truncated = True
            Exit Sub
        End If
        If MatchPattern(f.Name, patterns) Then
            filesScanned = filesScanned + 1
            Dim content As String
            On Error Resume Next
            content = ReadTextAuto(f.Path)
            On Error GoTo 0
            Dim lines() As String
            lines = Split(Replace(content, vbCrLf, vbLf), vbLf)
            Dim i As Long
            For i = 0 To UBound(lines)
                MaybePump
                Dim hit As Boolean
                If isRegex Then
                    hit = re.Test(lines(i))
                ElseIf caseSensitive Then
                    hit = InStr(lines(i), query) > 0
                Else
                    hit = InStr(1, lines(i), query, vbTextCompare) > 0
                End If
                If hit Then
                    If count >= maxResults Then
                        truncated = True
                        Exit Sub
                    End If
                    Dim d As Object
                    Set d = CreateObject("Scripting.Dictionary")
                    d("file") = RelPath(base, f.Path)
                    d("line") = i + 1
                    d("text") = Trim$(lines(i))
                    results.Add d
                    count = count + 1
                End If
            Next
        End If
    Next
    If truncated Then Exit Sub
    Dim subf As Object
    For Each subf In folder.SubFolders
        If truncated Then Exit Sub
        Dim nm As String
        nm = LCase$(subf.Name)
        If nm <> ".git" And nm <> ".svn" And nm <> ".hg" And nm <> "node_modules" Then
            SearchWalk fso, subf, patterns, re, query, isRegex, caseSensitive, maxResults, results, count, truncated, filesScanned, base
        End If
    Next
End Sub

' =====================================================================
' text helpers
' =====================================================================

Private Function NormalizeNewlines(ByVal s As String) As String
    NormalizeNewlines = Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function JoinLines(arr() As String, ByVal count As Long) As String
    If count <= 0 Then
        JoinLines = ""
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To count - 1)
    Dim i As Long
    For i = 0 To count - 1
        MaybePump
        parts(i) = arr(i)
    Next
    JoinLines = Join(parts, vbCrLf)
End Function

Private Function JoinLinesFrom(arr() As String, ByVal startIdx As Long) As String
    Dim n As Long
    n = (UBound(arr) + 1) - startIdx
    If n > 0 Then
        If arr(UBound(arr)) = "" Then n = n - 1
    End If
    If n <= 0 Then
        JoinLinesFrom = ""
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To n - 1)
    Dim i As Long
    For i = 0 To n - 1
        MaybePump
        parts(i) = arr(startIdx + i)
    Next
    JoinLinesFrom = Join(parts, vbCrLf)
End Function

' =====================================================================
' schema helpers
' =====================================================================

Private Function MakeTool(ByVal name As String, ByVal desc As String, ByVal schema As Object) As Object
    Dim t As Object
    Set t = CreateObject("Scripting.Dictionary")
    t("name") = name
    t("description") = desc
    Set t("inputSchema") = schema
    Set MakeTool = t
End Function

Private Function SchemaWith(ByVal props As Object, Optional ByVal required As Object) As Object
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s("type") = "object"
    If Not props Is Nothing Then Set s("properties") = props
    If Not required Is Nothing Then Set s("required") = required
    Set SchemaWith = s
End Function

Private Function Prop(ByVal t As String, ByVal desc As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("type") = t
    d("description") = desc
    Set Prop = d
End Function

Private Function EnumList(ParamArray values() As Variant) As Object
    Dim c As Object
    Set c = New Collection
    Dim i As Long
    For i = LBound(values) To UBound(values)
        c.Add CStr(values(i))
    Next
    Set EnumList = c
End Function

Private Function Req(ParamArray names() As Variant) As Object
    Dim c As Object
    Set c = New Collection
    Dim i As Long
    For i = LBound(names) To UBound(names)
        c.Add CStr(names(i))
    Next
    Set Req = c
End Function

' =====================================================================
' integration tools (offline-verified modules)
' =====================================================================

Private Function InferKind(ByVal name As String) As String
    Dim n As String
    n = LCase$(name)
    If Right$(n, 4) = ".cls" Then
        InferKind = "cls"
    ElseIf Right$(n, 4) = ".frm" Then
        InferKind = "frm"
    Else
        InferKind = "bas"
    End If
End Function

Private Function LintResultJson(ByVal name As String, ByVal kind As String, ByVal mode As String, ByVal text As String, Optional ByVal wantRefusalReason As Boolean = False, Optional ByVal baselineText As String = "") As String
    ' TASK-P1-093 (F2): this receipt carries the FULL diagnostics array (it has no cap), so
    ' ask the engine for UNLIMITED storage: DiagStored = DiagCount and the emitted array is
    ' byte-identical to before. errN/warnN now come from the engine''s own counters - same
    ' values, without the 27,005-iteration array scan (each iteration was a cross-module
    ' DiagSev BSTR call plus a pump point).
    Vb6Lint.SetStoreCap 0
    Vb6Lint.Lint text, kind
    Dim errN As Long, warnN As Long, i As Long
    errN = Vb6Lint.ErrorCount()
    warnN = Vb6Lint.WarningCount()
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = name
    root("kind") = kind
    root("mode") = mode
    root("ok") = (errN = 0)
    root("rejected") = (LCase$(mode) = "strict" And errN > 0)
    root("lintErrors") = errN
    root("lintWarnings") = warnN
    root("baselineHits") = Vb6Lint.BaselineHits()
    root("sha256") = Sha256.Sha256HexTextUtf8(text)
    root("defaultSource") = "edit"
    ' TASK-P1-121 item 8 (F-b): say out loud what this lint verdict IS and IS NOT.
    ' P4-100 measured four HARD compile errors (missing argument, undefined procedure,
    ' Exit Function inside a Sub, assignment to a module-level Const) on which ok stayed
    ' true. `ok` is DELIBERATELY left exactly as it was - (errN = 0) - and these new
    ' fields carry the honesty instead. Pure addition: no existing key or value moved.
    root("verdict") = "UNKNOWN"
    root("compilesVerified") = False
    Dim cannot As Object
    Set cannot = New Collection
    cannot.Add "compiler:missing_argument"
    cannot.Add "compiler:undefined_procedure"
    cannot.Add "compiler:wrong_exit_kind (Exit Function in Sub / Exit Sub in Function)"
    cannot.Add "compiler:assignment_to_constant"
    cannot.Add "compiler:type_mismatch"
    Set root("cannotDetect") = cannot
    root("ruleCount") = Vb6Lint.RuleCount()
    Dim arr As Object
    Set arr = New Collection
    ' TASK-P1-093 (F2): DiagStored, not DiagCount - with the uncapped setting above they
    ' are equal, and the guard keeps the loop safe if a cap is ever in force here.
    ' TASK-P1-126 (MCP-15): ONE pass over the linted text gives every diagnostic its
    ' STABLE identity - the enclosing procedure and the offset inside it - so a line
    ' number that moved with an insertion can no longer masquerade as a new error.
    Dim pmNames() As String, pmStarts() As Long
    If Len(text) > 0 Then Vb6Lint.LintProcMap text, pmNames, pmStarts
    For i = 0 To Vb6Lint.DiagStored() - 1
        MaybePump
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        Dim dLine As Long
        dLine = Vb6Lint.DiagLine(i)
        Dim dProc As String, dOff As Long
        dProc = ""
        dOff = 0
        If dLine >= 1 And dLine <= UBound(pmNames) + 1 Then
            dProc = pmNames(dLine - 1)
            If pmStarts(dLine - 1) > 0 Then dOff = dLine - pmStarts(dLine - 1) + 1
        End If
        d("ruleId") = Vb6Lint.DiagRule(i)
        d("severity") = Vb6Lint.DiagSev(i)
        d("editLine") = dLine
        d("message") = Vb6Lint.DiagMsg(i)
        d("proc") = dProc
        d("procOffset") = dOff
        d("fingerprint") = Vb6Lint.DiagRule(i) & "|" & Vb6Lint.DiagMsg(i) & "|" & dProc & "|" & CStr(dOff)
        d("stableKey") = Vb6Lint.DiagRule(i) & "|" & LCase$(Vb6Lint.DiagMsg(i)) & "|" & dProc
        arr.Add d
    Next
    Set root("diagnostics") = arr
    ' TASK-P1-126 (MCP-15): baseline_diff is ALWAYS present (discipline 98) so a caller
    ' can branch on it without a shape check; it is enabled=false without a baseline.
    Dim bdiff As Object
    If Len(baselineText) > 0 Then
        Set bdiff = BaselineDiffDict(baselineText, arr)
    Else
        Set bdiff = CreateObject("Scripting.Dictionary")
        bdiff("enabled") = False
        bdiff("matchedOn") = "stableKey (ruleId|message|proc)"
        bdiff("fingerprintField") = "fingerprint = ruleId|message|proc|procOffset"
        bdiff("baselineKeyCount") = 0
        Set bdiff("added") = New Collection
        Set bdiff("removed") = New Collection
        bdiff("addedCount") = 0
        bdiff("removedCount") = 0
        bdiff("unchangedCount") = 0
        bdiff("note") = "pass baseline=<the previous receipt text or its stableKey lines> to get added[]/removed[]"
    End If
    Set root("baseline_diff") = bdiff
    ' TASK-P1-110 (D-1107-03): check_component answered ok:false with NO structured reason, so a
    ' caller could not branch on the refusal without parsing prose. ADDITIVE and FAILURE-ONLY:
    ' the success receipt is byte-identical. validate_code keeps its old receipt (it calls this
    ' with the flag omitted = False); that sibling is LISTED in the report, not fixed here.
    If wantRefusalReason And (errN > 0) Then
        root("reason") = "lint_errors"
        root("suggestedAction") = "fix the severity=error diagnostics listed in 'diagnostics' and call again"
        root("refusalDetail") = "read-only tool: nothing was changed. The component's current IDE text has " & CStr(errN) & " severity=error diagnostic(s)."
    End If
    LintResultJson = JsonStringify(root)
End Function

' =====================================================================
' Contract v1.2.0 A batch (TASK-P1-020): alpha receipt honesty + beta guards.
' Nothing in this block changes an existing number or an existing field's
' meaning: it labels the coordinate space a number belongs to, reports whether
' the compiler actually refreshed its artifact, and refuses an edit that would
' silently split a line or a VB continuation.
' =====================================================================

' The line-feed normaliser the guard helpers use.
Private Function TextLines(ByVal text As String) As String()
    TextLines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Function

' =====================================================================
' TASK-P1-161 (R4): the pane line -> disk line map, ACCUMULATED per line.
'
' The mapping used to be ONE SCALAR per component: diskLine = ideLine + headerOffset,
' where headerOffset counted only the disk lines BEFORE the pane's first line. VBA6
' hides MORE of the file than that: a member attribute stored as its own
' `Attribute Foo.VB_UserMemId = -4` line sits inside the module body, is present in
' the .cls on disk, and NEVER appears in the code pane. Every pane line after such a
' line therefore got a disk number that was too small by one per hidden line, and a
' proc:/symbol: slice was numbered from a wrong base as well.
'
' The map below aligns the disk text with the pane text line by line and counts each
' hidden line AS IT IS PASSED, so the offset of line N is the SUM of the hidden lines
' above it. The alignment validates itself: a pane line that cannot be matched fixes
' resolved=False, and the receipt then says the number is NOT exact instead of
' publishing a scalar that only holds near the top of the file.
' =====================================================================

' True when a DISK line is a member attribute line VBA6 hides from the code pane:
' "Attribute <Name>.<Member> = <value>". The LEADING module attributes
' ("Attribute VB_Name = ...") belong to the header, not to this test; the dot before
' the '=' is what separates the two forms.
Private Function IsInlineAttributeLine(ByVal s As String) As Boolean
    IsInlineAttributeLine = False
    Dim t As String
    t = Trim$(s)
    If Len(t) < 11 Then Exit Function
    If LCase$(Left$(t, 10)) <> "attribute " Then Exit Function
    Dim eq As Long
    eq = InStr(t, "=")
    If eq = 0 Then Exit Function
    Dim dot As Long
    dot = InStr(t, ".")
    If dot > 11 And dot < eq Then IsInlineAttributeLine = True
End Function

' Line equality for the alignment. The RAW compare is tried first because the code pane
' expands TABs; CollapseWs is only paid on the lines that really differ.
Private Function LinesEqNorm(ByVal a As String, ByVal b As String) As Boolean
    If a = b Then
        LinesEqNorm = True
        Exit Function
    End If
    LinesEqNorm = (CollapseWs(a) = CollapseWs(b))
End Function

' Align pane lines to disk lines from a given header length. Fills map() (0-based,
' value = 1-based disk line, 0 = unmapped) and the two output collections; returns the
' number of DEVIATING pane lines (0 = the whole module aligned).
Private Function AlignPaneToDisk(ByRef dl() As String, ByVal nDisk As Long, _
                                 ByRef pl() As String, ByVal nPane As Long, _
                                 ByVal hdr As Long, ByRef map() As Long, _
                                 ByVal hidden As Object, ByVal dev As Object) As Long
    Dim di As Long, pi As Long, ndev As Long, k As Long
    Dim matched As Boolean, resynced As Boolean, guard As Long
    di = hdr
    pi = 0
    ndev = 0
    Do While pi <= nPane
        MaybePump
        If di > nDisk Then
            dev.Add pi + 1
            ndev = ndev + 1
            map(pi) = 0
            pi = pi + 1
        Else
            matched = False
            If LinesEqNorm(dl(di), pl(pi)) Then
                matched = True
            Else
                guard = 0
                Do While di <= nDisk And guard < 200
                    If IsInlineAttributeLine(dl(di)) Then
                        hidden.Add di + 1
                        di = di + 1
                        guard = guard + 1
                        If di <= nDisk Then
                            If LinesEqNorm(dl(di), pl(pi)) Then
                                matched = True
                                Exit Do
                            End If
                        End If
                    Else
                        Exit Do
                    End If
                Loop
            End If
            If matched Then
                map(pi) = di + 1
                di = di + 1
                pi = pi + 1
            Else
                ' bounded resync: an extra DISK line that is not an Attribute line
                resynced = False
                k = 1
                Do While k <= 4
                    If di + k <= nDisk Then
                        If LinesEqNorm(dl(di + k), pl(pi)) Then
                            resynced = True
                            Exit Do
                        End If
                    End If
                    k = k + 1
                Loop
                If resynced Then
                    di = di + k
                    map(pi) = di + 1
                    di = di + 1
                    pi = pi + 1
                Else
                    dev.Add pi + 1
                    ndev = ndev + 1
                    map(pi) = 0
                    pi = pi + 1
                End If
            End If
        End If
    Loop
    AlignPaneToDisk = ndev
End Function

' The map for one component (not cached here - use PaneDiskMapCached).
Private Function PaneDiskMap(ByVal name As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("component") = name
    d("resolved") = False
    d("headerLines") = 0
    d("hiddenCount") = 0
    d("paneLineCount") = 0
    d("diskLineCount") = 0
    d("deviationCount") = 0
    d("method") = "cumulative pane->disk alignment: diskLine = ideLine + headerLines + every hidden inline Attribute line ABOVE that line"
    d("basis") = "no map was computed (the component could not be resolved with a code module)"
    Dim hidden As Object, dev As Object
    Set hidden = New Collection
    Set dev = New Collection
    Set d("hiddenDiskLines") = hidden
    Set d("deviationPaneLines") = dev
    Dim mapArr() As Long
    Set PaneDiskMap = d
    On Error Resume Next
    Dim path As String
    path = DiskPathOf(name, "")
    If Len(path) = 0 Then Err.Clear: Exit Function
    If Dir$(path) = "" Then Err.Clear: Exit Function
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Err.Clear: Exit Function
    If comp.CodeModule Is Nothing Then Err.Clear: Exit Function
    Dim ideLines As Long
    ideLines = comp.CodeModule.CountOfLines
    d("paneLineCount") = ideLines
    Dim pane As String
    If ideLines > 0 Then pane = comp.CodeModule.Lines(1, ideLines)
    Dim disk As String
    disk = Build.ReadLogAnsi(path)
    If Len(disk) = 0 Then Err.Clear: Exit Function
    Dim dl() As String, pl() As String
    dl = Split(Replace(Replace(disk, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    pl = Split(Replace(Replace(pane, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim nDisk As Long, nPane As Long
    nDisk = UBound(dl)
    If nDisk > 0 Then
        If Len(dl(nDisk)) = 0 Then nDisk = nDisk - 1
    End If
    If ideLines > 0 Then nPane = UBound(pl) Else nPane = -1
    d("diskLineCount") = nDisk + 1
    If nPane < 0 Then
        d("resolved") = True
        d("headerLines") = nDisk + 1
        d("basis") = "the code pane is empty, so the whole file is header"
        Err.Clear
        Exit Function
    End If
    ' candidate header positions: the disk lines equal to the pane's FIRST line.
    Dim cand(1 To 8) As Long
    Dim ncand As Long, i As Long
    ncand = 0
    For i = 0 To nDisk
        If dl(i) = pl(0) Then
            ncand = ncand + 1
            If ncand > 8 Then Exit For
            cand(ncand) = i
        End If
    Next
    If ncand = 0 Then
        ' the raw text differs (TAB expansion, a rewritten line, ...): retry normalized
        For i = 0 To nDisk
            MaybePump
            If i > 2000 Then Exit For
            If LinesEqNorm(dl(i), pl(0)) Then
                ncand = ncand + 1
                If ncand > 8 Then Exit For
                cand(ncand) = i
            End If
        Next
    End If
    If ncand = 0 Then
        d("basis") = "the pane's first line was not found in the file on disk, so no mapping could be derived (the offset is NOT exact)"
        Err.Clear
        Exit Function
    End If
    ' use the FIRST candidate that aligns the WHOLE module with zero deviations
    Dim tryNo As Long, devCount As Long, bestDev As Long
    bestDev = -1
    For tryNo = 1 To ncand
        Set hidden = New Collection
        Set dev = New Collection
        ReDim mapArr(0 To nPane)
        devCount = AlignPaneToDisk(dl, nDisk, pl, nPane, cand(tryNo), mapArr, hidden, dev)
        If bestDev < 0 Or devCount < bestDev Then
            bestDev = devCount
            Set d("hiddenDiskLines") = hidden
            Set d("deviationPaneLines") = dev
            d("headerLines") = cand(tryNo)
            d("hiddenCount") = hidden.Count
            d("deviationCount") = dev.Count
            d("ideToDisk") = mapArr
        End If
        If devCount = 0 Then Exit For
    Next
    d("resolved") = (bestDev = 0)
    If bestDev = 0 Then
        d("basis") = "diskLine = ideLine + headerLines(" & CStr(d("headerLines")) & ") + hiddenInlineAttributeLinesAbove(" & CStr(d("hiddenCount")) & " total); accumulated per line from the .cls on disk and validated by a whole-module alignment with 0 deviating pane lines"
    Else
        d("basis") = "NOT exact: the .cls on disk and the code pane could not be aligned (the best pass still had " & CStr(bestDev) & " deviating pane line(s)); the scalar headerOffset is reported instead and MUST NOT be trusted below the first hidden line"
    End If
    Err.Clear
End Function

' Per-top-level-call cache (m_paneDiskMaps is reset in CallTool).
Private Function PaneDiskMapCached(ByVal name As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("component") = name
    d("resolved") = False
    d("headerLines") = 0
    d("hiddenCount") = 0
    d("paneLineCount") = 0
    d("diskLineCount") = 0
    d("deviationCount") = 0
    d("method") = "cumulative pane->disk alignment"
    d("basis") = "no map was computed (the component could not be resolved with a code module)"
    Set d("hiddenDiskLines") = New Collection
    Set d("deviationPaneLines") = New Collection
    Set PaneDiskMapCached = d
    On Error Resume Next
    If m_paneDiskMaps Is Nothing Then Set m_paneDiskMaps = CreateObject("Scripting.Dictionary")
    If m_paneDiskMaps.Exists(name) Then
        Set PaneDiskMapCached = m_paneDiskMaps(name)
        Err.Clear
        Exit Function
    End If
    Dim m As Object
    Set m = PaneDiskMap(name)
    Set m_paneDiskMaps(name) = m
    Set PaneDiskMapCached = m
    Err.Clear
End Function

' =====================================================================
' TASK-P1-164 (R18-2): the post-operation assertion for tools that may not write text.
' Same contract as the editor's expected_sha256 - a mismatch is ok:false /
' reason:paneTextMismatch WITH the differing lines, and "could not read" is its own
' reported state; neither is ever a silent pass.
' =====================================================================

' The code-pane text of a component. got=False means it could not be read.
Private Function PaneTextOfName(ByVal name As String, ByRef got As Boolean) As String
    got = False
    On Error Resume Next
    PaneTextOfName = ""
    If Len(name) = 0 Then Exit Function
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    If comp.CodeModule Is Nothing Then Exit Function
    Dim n As Long
    n = comp.CodeModule.CountOfLines
    If n > 0 Then PaneTextOfName = comp.CodeModule.Lines(1, n)
    got = True
    Err.Clear
End Function

' The component of the ACTIVE code pane ("" when there is none).
Private Function PaneCompOfActive() As String
    PaneCompOfActive = ""
    On Error Resume Next
    Dim cp As VBIDE.CodePane
    Set cp = m_vbe.ActiveCodePane
    If cp Is Nothing Then Exit Function
    PaneCompOfActive = cp.CodeModule.Name
    Err.Clear
End Function

' The component whose code pane IS this window. Two EXACT channels, in order:
'   1. hwnd  - the window handle reported by list_windows (Window.HWnd) equals the
'              code pane's own window handle. Measured 2026-10-13: this channel did NOT
'              match on the live IDE, so it is tried first and reported when it hits.
'   2. caption_exact - Window.Caption and CodePane.Window.Caption are the SAME string;
'              equality (never a substring) is what keeps this from being a guess.
' via reports which channel matched, so the receipt never implies more than was checked.
Private Function PaneCompOfWindowTarget(ByVal hwnd As Long, ByVal caption As String, ByRef via As String) As String
    PaneCompOfWindowTarget = ""
    via = ""
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If p Is Nothing Then Exit Function
    Dim i As Long
    Dim c As VBIDE.VBComponent
    Dim cp As VBIDE.CodePane
    If hwnd <> 0 Then
        For i = 1 To p.VBComponents.Count
            Set c = p.VBComponents.Item(i)
            If Not c Is Nothing Then
                Set cp = Nothing
                Set cp = c.CodeModule.CodePane
                If Not cp Is Nothing Then
                    If cp.Window.HWnd = hwnd Then
                        PaneCompOfWindowTarget = c.Name
                        via = "hwnd"
                        Err.Clear
                        Exit Function
                    End If
                End If
            End If
            Err.Clear
        Next
    End If
    If Len(caption) > 0 Then
        For i = 1 To p.VBComponents.Count
            Set c = p.VBComponents.Item(i)
            If Not c Is Nothing Then
                Set cp = Nothing
                Set cp = c.CodeModule.CodePane
                If Not cp Is Nothing Then
                    If cp.Window.Caption = caption Then
                        PaneCompOfWindowTarget = c.Name
                        via = "caption_exact"
                        Err.Clear
                        Exit Function
                    End If
                End If
            End If
            Err.Clear
        Next
    End If
    Err.Clear
End Function

' The first N differing lines of two pane texts.
Private Function PaneDiffLinesOf(ByVal before As String, ByVal after As String, ByVal maxN As Long) As Object
    Dim c As Object
    Set c = New Collection
    Dim bl() As String, al() As String
    bl = Split(Replace(Replace(before, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    al = Split(Replace(Replace(after, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim n As Long, i As Long
    n = UBound(bl)
    If UBound(al) > n Then n = UBound(al)
    For i = 0 To n
        Dim b1 As String, a1 As String
        b1 = ""
        a1 = ""
        If i <= UBound(bl) Then b1 = bl(i)
        If i <= UBound(al) Then a1 = al(i)
        If b1 <> a1 Then
            Dim e As Object
            Set e = CreateObject("Scripting.Dictionary")
            e("line") = i + 1
            e("before") = Left$(b1, 200)
            e("after") = Left$(a1, 200)
            c.Add e
            If c.Count >= maxN Then Exit For
        End If
    Next
    Set PaneDiffLinesOf = c
End Function

' Apply the assertion to a receipt. on_=False means the BEFORE text was not obtained;
' that is reported as notApplicable and never as a pass.
Private Sub PaneAssertFinish(ByVal root As Object, ByVal name As String, ByVal before As String, _
                             ByVal on_ As Boolean, ByVal label As String)
    If Not on_ Then
        root("paneAssertion") = "notApplicable"
        root("paneAssertionNote") = "the pane text could not be read BEFORE the operation, so nothing is claimed about it (" & label & ")"
        Exit Sub
    End If
    Dim got As Boolean
    Dim after As String
    after = PaneTextOfName(name, got)
    If Not got Then
        root("paneAssertion") = "notApplicable"
        root("paneAssertionNote") = "the pane text of " & name & " could not be read AFTER the operation, so no claim is made (" & label & ")"
        Exit Sub
    End If
    root("paneAssertion") = "compared"
    root("paneComponent") = name
    root("paneSha256Before") = Sha256.Sha256HexTextUtf8(before)
    root("paneSha256After") = Sha256.Sha256HexTextUtf8(after)
    If before = after Then
        root("paneTextUnchanged") = True
        Exit Sub
    End If
    root("paneTextUnchanged") = False
    root("ok") = False
    root("reason") = "paneTextMismatch"
    Set root("paneDiffLines") = PaneDiffLinesOf(before, after, 5)
    root("paneMismatchNote") = "this tool is not supposed to write any text, but the code pane of " & name & " changed during the call. Nothing is claimed to have been written correctly: re-read the component and compare paneSha256Before/paneSha256After before trusting anything else in this receipt."
End Sub


' coords{headerOffset, ideLine, diskLine, default} with headerOffset = diskLine - ideLine.
' 'default' names the line space the RESPONSE speaks about. It is per tool on
' purpose: get_component_code PREPENDS the file header, so its emitted numbers are
' DISK lines, while the editing tools speak IDE lines. CONTRACTS.md invariant 4
' moves the default to "ide" only when the B+C batch returns the pure code text.
Private Function CoordsDict(ByVal ideLine As Long, ByVal headerOffset As Long, ByVal defaultCoord As String, ByVal resolved As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("headerOffset") = headerOffset
    d("ideLine") = ideLine
    d("diskLine") = ideLine + headerOffset
    d("default") = defaultCoord
    d("resolved") = resolved
    Set CoordsDict = d
End Function

' headerOffset = disk physical lines - IDE CodeModule lines. resolved=False when
' the component has no file on disk (the offset is then 0 by convention and MUST
' NOT be trusted) or the code module cannot be read.
' headerOffset = the number of DISK lines that precede the component's first
' code line, i.e. exactly what get_component_code prepends. Preferring the LOCATED
' header over (disk lines - pane lines) means the offset stays exact while the
' pane is dirty (an unsaved edit changes the line counts on one side only).
' resolved=False means the pane's first line is not on disk: the offset cannot be
' trusted and the caller must not use it.
Private Sub ComponentOffsetOf(ByVal name As String, ByRef off As Long, ByRef resolved As Boolean)
    On Error Resume Next
    off = 0
    resolved = False
    Dim path As String
    path = DiskPathOf(name, "")
    If Len(path) = 0 Then Exit Sub
    If Dir$(path) = "" Then Exit Sub
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Sub
    If comp.CodeModule Is Nothing Then Exit Sub
    Dim ideLines As Long
    ideLines = comp.CodeModule.CountOfLines
    Dim dl As Long
    dl = HeaderOffsetForPath(path, 0)
    If ideLines = 0 Then
        off = dl
        resolved = True
        Exit Sub
    End If
    Dim pane As String
    pane = comp.CodeModule.Lines(1, ideLines)
    Dim h As String
    h = ComponentHeaderOf(name, pane)
    If Len(h) > 0 Then
        off = UBound(Split(h, vbCrLf))
        resolved = True
        Exit Sub
    End If
    Dim dlines() As String
    Dim plines() As String
    dlines = Split(Replace(Build.ReadLogAnsi(path), vbCrLf, vbLf), vbLf)
    plines = Split(Replace(pane, vbCrLf, vbLf), vbLf)
    If dlines(0) = plines(0) Then
        off = 0
        resolved = True
        Exit Sub
    End If
    off = dl - ideLines
    resolved = False
End Sub

Private Function CoordsDictOfName(ByVal name As String, ByVal ideLine As Long, ByVal defaultCoord As String) As Object
    Dim off As Long, resolved As Boolean
    ComponentOffsetOf name, off, resolved
    Dim d As Object
    Set d = CoordsDict(ideLine, off, defaultCoord, resolved)
    ' TASK-P1-161 (R4): diskLine is ACCUMULATED per line, not a constant offset, and the
    ' receipt carries the mapping basis so a caller can audit the number instead of
    ' trusting one scalar (discipline 39: the fallback is labelled, never silent).
    Dim em As Object
    Set em = PaneDiskMapCached(name)
    d("mappingMethod") = em("method")
    d("mappingBasis") = em("basis")
    d("mappingResolved") = em("resolved")
    d("hiddenInlineAttributeLines") = em("hiddenCount")
    d("headerLines") = em("headerLines")
    d("diskLineExact") = False
    If CBool(em("resolved")) Then
        Dim exact As Long
        exact = ExactDiskLineOf(em, ideLine, 0)
        If exact > 0 Then
            d("diskLine") = exact
            d("diskLineExact") = True
            d("hiddenInlineAttributeLinesBefore") = exact - ideLine - CLng(em("headerLines"))
        End If
    End If
    Set CoordsDictOfName = d
End Function

' TASK-P1-161 (R4): the exact disk line of a pane line, or 0 when the map has no
' trustworthy answer for it (the caller then keeps its own scalar fallback).
Private Function ExactDiskLineOf(ByVal em As Object, ByVal ideLine As Long, ByVal fallbackNo As Long) As Long
    ExactDiskLineOf = fallbackNo
    On Error Resume Next
    If em Is Nothing Then Exit Function
    If Not CBool(em("resolved")) Then Exit Function
    Dim arr() As Long
    arr = em("ideToDisk")
    If ideLine < 1 Then Exit Function
    If ideLine > UBound(arr) + 1 Then Exit Function
    If arr(ideLine - 1) > 0 Then ExactDiskLineOf = arr(ideLine - 1)
    Err.Clear
End Function

Private Function CompLineCount(ByVal name As String) As Long
    On Error Resume Next
    CompLineCount = 0
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    If comp.CodeModule Is Nothing Then Exit Function
    CompLineCount = comp.CodeModule.CountOfLines
End Function

Private Function KindOfName(ByVal name As String) As String
    On Error Resume Next
    KindOfName = InferKind(name)
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    KindOfName = KindOfComponent(comp)
End Function

Private Function SyncWarningText(ByVal ideTotalLines As Long, ByVal textTotalLines As Long) As String
    If ideTotalLines = textTotalLines Then Exit Function
    SyncWarningText = "the IDE reports " & CStr(ideTotalLines) & " lines after the write but the written text has " & CStr(textTotalLines) & "; the pane and the text disagree - re-read with get_component_code and compare sha256Before/sha256After before trusting newTotalLines"
End Function

' alpha-5: prove whether a hash is the DISK file's bytes or the code pane's text.
Private Sub AddPaneDiskFields(ByVal root As Object, ByVal name As String, ByVal paneText As String)
    On Error Resume Next
    root("paneSha256") = Sha256.Sha256HexTextUtf8(paneText)
    root("diskSha256") = ComponentFileSha(name)
    root("onDisk") = DiskMatchesPane(name, paneText)
End Sub

' beta-3: run the E101-E116 + E107 + X/W rule engine on the edit RESULT. The A
' batch is warn-only by contract; the diagnostics ride along in the receipt so the
' caller can see what the write would introduce.
' TASK-P1-051: modules at or below this line count keep the ORIGINAL whole-module
' lint (bit-identical receipt). Above it, lint targets the edited procedure only.
' Configurable via the VB6MCP_LINT_REGION_THRESHOLD environment variable.
Private Function LintRegionThreshold() As Long
    Dim s As String, v As Long
    Const DEF As Long = 5000
    v = DEF
    s = Trim$(Environ$("VB6MCP_LINT_REGION_THRESHOLD"))
    If Len(s) > 0 Then
        On Error Resume Next
        v = CLng(s)
        If Err.Number <> 0 Then v = DEF
        On Error GoTo 0
    End If
    If v < 1 Then v = DEF
    LintRegionThreshold = v
End Function

' TASK-P1-086 (F1): the diagnostics ARRAY a receipt may carry is capped here. 300 is
' chosen from MEASUREMENT, not taste: an ordinary module in this repo produces fewer
' than 50 diagnostics, while the pathological module-level fixture (27,000 assignment
' lines, no procedure) produced exactly 27,000 and a 2,690,288-byte receipt (P1-074).
' 300 is ~6x the ordinary p99 and keeps the worst receipt in the tens of KB; the card's
' suggested band was 200-500. Nothing is lost silently: diagnosticsTotal carries the
' FULL count, lintTruncated says whether the array was cut, lintDiagnosticCap names the
' cap in force, and lintSummary still counts every diagnostic (StrictRefusal reads
' lintSummary("errors"), so validate:'strict' is NOT weakened by the cap).
' Configurable via VB6MCP_LINT_MAX_DIAG (mirrors VB6MCP_LINT_REGION_THRESHOLD).
Private Function LintMaxDiag() As Long
    Dim s As String, v As Long
    Const DEF As Long = 300
    v = DEF
    s = Trim$(Environ$("VB6MCP_LINT_MAX_DIAG"))
    If Len(s) > 0 Then
        On Error Resume Next
        v = CLng(s)
        If Err.Number <> 0 Then v = DEF
        On Error GoTo 0
    End If
    If v < 1 Then v = DEF
    LintMaxDiag = v
End Function

' TASK-P1-094: the anchor batch path WAS O(ops x module lines), because
' EditTxn.ApplyLineOps rebuilt the WHOLE text once per op (and called CountLines twice
' per op). MEASURED on the 27,000-line fixture BEFORE this guard existed: s=1 -> 300 ms,
' s=50 -> 4,605 ms (92 ms/op), s=500 -> the add-in closed the connection (the pre-fix
' build died at 33 s on the same input). A batch is therefore REFUSED UP FRONT - before
' any work - when ops x lines exceeds this budget, with a structured reason and the
' largest op count that WOULD fit, instead of turning the limit into a hang or a dropped
' connection (TASK-P1-094 criteria 2 and 4).
'
' TASK-P1-097 UPDATE - the numbers above describe the cost the guard was introduced for;
' they are NOT the current cost. ApplyLineOps is now single-pass (one line array, one
' Join) and the anchor path resolves off one batch index, so ops x lines no longer models
' the work. RE-MEASURED on the same 27,011-line fixture with the default budget in force:
' s=1 278 ms, s=10 356 ms, s=50 444 ms, s=74 537 ms; the marginal cost from s=50 to s=74
' is ~1.0 ms/op (was 106 ms/op), i.e. the batch engine is O(lines + ops) now. The default
' is therefore kept as a CONSERVATIVE SAFETY NET rather than as the O(ops x lines) guard it
' was introduced as: 2,000,000 op-lines admits at most 74 ops at 27k lines (~0.5 s
' measured) and ~1994 ops at 1k lines (s=500 there measured 767 ms - sub-second). Raising
' the default is a product decision, not a fix: with the budget lifted to 40,000,000 the
' same s=500 request at 27k COMPLETES in 7.5 s (no dropped connection), but what is left is
' the preview_edit DIFF TEXT plus a 1.18 MB JSON response, not the batch engine
' (REP-TASK-P1-097 section 4).
' Configurable via VB6MCP_ANCHOR_MAX_WORK (same style as VB6MCP_LINT_MAX_DIAG).
Private Function AnchorMaxWork() As Double
    Dim s As String, v As Double
    Const DEF As Double = 2000000#
    v = DEF
    s = Trim$(Environ$("VB6MCP_ANCHOR_MAX_WORK"))
    If Len(s) > 0 Then
        On Error Resume Next
        v = CDbl(s)
        If Err.Number <> 0 Then v = DEF
        On Error GoTo 0
    End If
    If v < 1000# Then v = 1000#
    AnchorMaxWork = v
End Function

' "" when the batch fits, else the STRUCTURED rejection receipt (nothing is touched).
Private Function AnchorBatchTooLarge(ByVal toolName As String, ByVal name As String, ByVal opN As Long, ByVal lineN As Long) As String
    AnchorBatchTooLarge = ""
    If opN <= 0 Or lineN <= 0 Then Exit Function
    Dim w As Double, cap As Double, sug As Long
    w = CDbl(opN) * CDbl(lineN)
    cap = AnchorMaxWork()
    If w <= cap Then Exit Function
    sug = CLng(cap / CDbl(lineN))
    If sug < 1 Then sug = 1
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "batch_too_large"
    d("component") = name
    d("tool") = toolName
    d("ops") = opN
    d("ideTotalLines") = lineN
    d("workEstimate") = CLng(w)
    d("workBudget") = CLng(cap)
    d("suggestedMaxOps") = sug
    d("note") = "the anchor batch cost is O(ops x lines): this request needs " & CStr(CLng(w)) & " op-lines. Send at most " & CStr(sug) & " ops, split the batch, or raise VB6MCP_ANCHOR_MAX_WORK (default " & CStr(CLng(cap)) & "). The refusal happens BEFORE any work - nothing was written and the pane is untouched."
    AnchorBatchTooLarge = JsonStringify(d)
End Function

' TASK-P1-051: which IDE line identifies the procedure a large-module lint targets.
' 0 = no region (lint the whole module); "set" replaces the entire module, so there
' is no single target procedure.
Private Function LintRegionTarget(ByVal op As String, ByVal adjStart As Long, ByVal countOfLines As Long) As Long
    LintRegionTarget = 0
    Select Case op
        Case "replace", "delete", "insert"
            LintRegionTarget = adjStart
        Case "append"
            LintRegionTarget = countOfLines
        Case Else
            LintRegionTarget = 0
    End Select
    If LintRegionTarget < 1 Then LintRegionTarget = 0
End Function

' TASK-P1-054: the same region policy as replace_code, for every other edit tool that
' lints before writing. sizeLines is the CURRENT pane line count (the threshold test is
' about module size), targetLine is the IDE line the tool is about to touch (0 = the
' tool has no single target line, in which case the whole module is linted and the
' receipt claims NO region scope).
Private Function LintPackForEdit(ByVal text As String, ByVal kind As String, ByVal sizeLines As Long, ByVal targetLine As Long) As Object
    Dim t As Long
    If sizeLines > LintRegionThreshold() Then
        If targetLine > 0 Then t = targetLine
    End If
    Set LintPackForEdit = LintPackScoped(text, kind, t)
End Function

' TASK-P1-054: best-effort physical line of a procedure declaration, used as the region
' target by inject_trace / fault_inject (they know the NAME but not a line). 0 = not
' found, and the caller then lints the whole module rather than guessing.
Private Function ProcHeaderLineOf(ByVal text As String, ByVal procName As String) As Long
    ProcHeaderLineOf = 0
    Dim nm As String, lines() As String, i As Long, s As String, t As String, p As Long
    nm = LCase$(Trim$(procName))
    If Len(nm) = 0 Then Exit Function
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        MaybePump
        s = LCase$(Trim$(lines(i)))
        Do
            t = s
            If Left$(s, 7) = "public " Then s = Mid$(s, 8)
            If Left$(s, 8) = "private " Then s = Mid$(s, 9)
            If Left$(s, 7) = "friend " Then s = Mid$(s, 8)
            If Left$(s, 7) = "static " Then s = Mid$(s, 8)
            If t = s Then Exit Do
        Loop
        If Left$(s, 4) = "sub " Then
            s = Mid$(s, 5)
        ElseIf Left$(s, 9) = "function " Then
            s = Mid$(s, 10)
        ElseIf Left$(s, 9) = "property " Then
            s = Mid$(s, 10)
            p = InStr(s, " ")
            If p > 0 Then s = Mid$(s, p + 1)
        Else
            GoTo NextLine054
        End If
        p = InStr(s, "(")
        If p > 0 Then s = Left$(s, p - 1)
        If Trim$(s) = nm Then
            ProcHeaderLineOf = i + 1
            Exit Function
        End If
NextLine054:
    Next i
End Function

Private Function LintRegionDict(ByVal procName As String, ByVal r0 As Long, ByVal r1 As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("proc") = procName
    d("startLine") = r0
    d("endLine") = r1
    Set LintRegionDict = d
End Function

' beta-3 / TASK-P1-051: run the E101-E116 + E107 + X/W rule engine on the edit
' RESULT. The A batch is warn-only by contract; the diagnostics ride along in the
' receipt so the caller can see what the write would introduce. targetLine > 0 asks
' for REGION scope (lint only the procedure containing that IDE line); when no
' procedure contains the line the WHOLE module is linted and no region scope is
' claimed. Region mode is only entered for modules above LintRegionThreshold().
Private Function LintPackScoped(ByVal text As String, ByVal kind As String, ByVal targetLine As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim scoped As Boolean
    Dim procName As String
    Dim r0 As Long, r1 As Long
    ' TASK-P1-093 (F2): this receipt can only ever SHOW LintMaxDiag() diagnostics, so tell
    ' the engine to materialise at most that many. It still counts every diagnostic, so
    ' lintSummary / diagnosticsTotal / lintTruncated / strict refusal are unchanged - but a
    ' 27,005-diagnostic module no longer pays 27,005 UDT+BSTR writes for entries the receipt
    ' can never carry. Single source for the cap = LintMaxDiag() (vb6 defaults to 0 = store
    ' everything, so a caller that does not opt in behaves exactly as before).
    Dim capEarly As Long
    capEarly = LintMaxDiag()
    Vb6Lint.SetStoreCap capEarly
    On Error Resume Next
    scoped = False
    procName = ""
    r0 = 0
    r1 = 0
    If targetLine > 0 Then
        scoped = Vb6Lint.LintRegionAt(text, kind, targetLine, procName, r0, r1)
    Else
        Vb6Lint.Lint text, kind
    End If
    Dim errN As Long, warnN As Long, i As Long
    Dim diagN As Long, capN As Long, keptN As Long
    ' TASK-P1-093 (F2): both numbers come from the engine counters, not from a full-array
    ' scan (the old loop walked all 27,005 entries from this class into Vb6Lint).
    errN = Vb6Lint.ErrorCount()
    warnN = Vb6Lint.WarningCount()
    diagN = Vb6Lint.DiagCount()
    ' TASK-P1-086 (F1): cap the diagnostics ARRAY (default 300 -> see LintMaxDiag) so a
    ' module-level edit in a 27,000-line module cannot inflate the receipt to MBs.
    capN = LintMaxDiag()
    keptN = Vb6Lint.DiagStored()
    If keptN > diagN Then keptN = diagN
    If keptN > capN Then keptN = capN
    Dim arr As Object
    Set arr = New Collection
    For i = 0 To keptN - 1
        Dim it As Object
        Set it = CreateObject("Scripting.Dictionary")
        it("ruleId") = Vb6Lint.DiagRule(i)
        it("severity") = Vb6Lint.DiagSev(i)
        it("editLine") = Vb6Lint.DiagLine(i)
        it("message") = Vb6Lint.DiagMsg(i)
        arr.Add it
    Next
    Set d("diagnostics") = arr
    ' TASK-P1-086 (F1): the full count, the truncation flag and the cap in force travel
    ' with the (possibly shortened) array. lintSummary above counts ALL diagnostics.
    d("diagnosticsTotal") = diagN
    d("lintTruncated") = (diagN > capN)
    d("lintDiagnosticCap") = capN
    Dim ls As Object
    Set ls = CreateObject("Scripting.Dictionary")
    ls("errors") = errN
    ls("warnings") = warnN
    Set d("lintSummary") = ls
    d("lintMode") = "warn"
    If scoped Then
        d("lintScope") = "region"
        Set d("lintRegion") = LintRegionDict(procName, r0, r1)
        d("lintNote") = "lint is this project's own rule engine, not the VB6 compiler; in the A batch it is warn-only and never blocks the write. This module is larger than " & CStr(LintRegionThreshold()) & " lines, so lintSCOPE is the edited procedure only: diagnostics OUTSIDE " & CStr(r0) & ".." & CStr(r1) & " are NOT reported here (validate:'strict' can therefore only be tripped by an in-region error). For a whole-module lint use verify_scope."
    ElseIf targetLine > 0 Then
        ' TASK-P1-086 (F2): the caller ASKED for a region (targetLine > 0) but no procedure
        ' contains that line, so LintRegionAt returned False and the WHOLE module was
        ' linted. That is the expensive path (P1-074: 0.14 ms/line, 3,797 ms for 27,000
        ' lines) and it used to be completely silent - the receipt looked exactly like a
        ' region pass. Label the fallback so the cost is attributable; lintNote keeps its
        ' existing value, the label lives in the two new keys.
        d("lintScope") = "whole"
        d("lintFallback") = "no_procedure_contains_line"
        d("lintNote") = "lint is this project's own rule engine, not the VB6 compiler; in the A batch it is warn-only and never blocks the write"
    Else
        d("lintNote") = "lint is this project's own rule engine, not the VB6 compiler; in the A batch it is warn-only and never blocks the write"
    End If
    Set LintPackScoped = d
End Function

' beta-3: the whole-module entry point (unchanged behaviour); kept because add_code /
' write_file / write_source / verify paths call it.
Private Function LintPack(ByVal text As String, ByVal kind As String) As Object
    Set LintPack = LintPackScoped(text, kind, 0)
End Function

Private Sub CopyLintPack(ByVal root As Object, ByVal pack As Object)
    On Error Resume Next
    If pack Is Nothing Then Exit Sub
    Set root("diagnostics") = pack("diagnostics")
    Set root("lintSummary") = pack("lintSummary")
    root("lintMode") = pack("lintMode")
    root("lintNote") = pack("lintNote")
    ' TASK-P1-051: these two keys exist ONLY when the lint really was region-scoped,
    ' so a small-module receipt is bit-identical to before.
    If pack.Exists("lintScope") Then root("lintScope") = pack("lintScope")
    If pack.Exists("lintRegion") Then Set root("lintRegion") = pack("lintRegion")
    ' TASK-P1-086 (F1/F2): the cap counters and the fallback label ride along with every
    ' lint-carrying receipt. Purely additive: the four keys copied above keep their exact
    ' values, and none of these four existed before this card.
    If pack.Exists("diagnosticsTotal") Then root("diagnosticsTotal") = pack("diagnosticsTotal")
    If pack.Exists("lintTruncated") Then root("lintTruncated") = pack("lintTruncated")
    If pack.Exists("lintDiagnosticCap") Then root("lintDiagnosticCap") = pack("lintDiagnosticCap")
    If pack.Exists("lintFallback") Then root("lintFallback") = pack("lintFallback")
End Sub

' beta-2: a logical line continued with "_" must not be cut in two. True when the
' 1-based lineNo is the TAIL of a logical line, i.e. the physical line before it
' ends with a VB continuation (underscore preceded by a space or a tab).
Private Function IsContinuationTail(ByRef lines() As String, ByVal lineNo As Long) As Boolean
    On Error Resume Next
    IsContinuationTail = False
    If lineNo <= 1 Then Exit Function
    Dim idx As Long
    idx = lineNo - 2
    If idx < 0 Then Exit Function
    If idx > UBound(lines) Then Exit Function
    Dim prev As String
    prev = RTrim$(Vb6Lint.StripCode(lines(idx)))
    If Len(prev) < 2 Then Exit Function
    If Right$(prev, 1) <> "_" Then Exit Function
    Dim pc As String
    pc = Mid$(prev, Len(prev) - 1, 1)
    If pc = " " Or pc = vbTab Then IsContinuationTail = True
End Function

Private Function RangeDict(ByVal operation As String, ByVal startLine As Long, ByVal endLine As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("operation") = operation
    d("start_line") = startLine
    d("end_line") = endLine
    Set RangeDict = d
End Function

Private Function AppliedRangeDict(ByVal startLine As Long, ByVal endLine As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("start_line") = startLine
    d("end_line") = endLine
    Set AppliedRangeDict = d
End Function

' beta-1: a structured refusal. The IDE is never touched on this path.
Private Function GuardReject(ByVal component As String, ByVal operation As String, ByVal reason As String, ByVal detail As String, ByVal startLine As Long, ByVal endLine As Long, ByVal countOfLines As Long) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("component") = component
    Set d("requestedRange") = RangeDict(operation, startLine, endLine)
    If reason = "range_beyond_end" Then d("countOfLines") = countOfLines
    ' TASK-P1-020 rework (contract ruling 4): a failure carries line numbers too, so it
    ' must carry coords as well - otherwise requestedRange / countOfLines cannot be
    ' interpreted. The edit tools take IDE (code pane) lines; get_component_code's
    ' NUMBERED output is DISK lines, so the hint spells the conversion out (a verifier
    ' mixed the two spaces and read a correct acceptance as a missed rejection).
    Set d("coords") = CoordsDictOfName(component, startLine, "ide")
    d("lineSpace") = "ide"
    d("lineSpaceHint") = "requestedRange / countOfLines are IDE code-pane lines. get_component_code's numbered output is DISK lines: ideLine = diskLine - coords.headerOffset"
    d("detail") = detail
    d("note") = "the component was NOT modified"
    GuardReject = JsonStringify(d)
End Function

' ===================== TASK-P1-126 helpers =====================
' (a) the PRE-WRITE pairing gate shared by replace_code / add_code / apply_edits /
'     edit_by_anchor;
' (b) the procedure first/last line echo of edit_by_anchor;
' (c) the machine-readable refusal for a missing expected_sha256 on a line-number
'     edit;
' (d) baseline_diff for validate_code / check_component.
' ===============================================================
Private Function PreWritePairingRefusal(ByVal toolName As String, ByVal component As String, _
                                        ByVal blockText As String, ByVal where As String) As String
    PreWritePairingRefusal = ""
    If Len(blockText) = 0 Then Exit Function
    Dim nh As Long, ne As Long, extra As Long
    Dim fixed As String
    Dim why As String
    why = Vb6Lint.BlockPairingProblem(blockText, nh, ne, extra, fixed)
    If Len(why) = 0 Then Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = why
    d("tool") = toolName
    d("component") = component
    d("where") = where
    d("writePerformed") = False
    d("note") = "refused BEFORE any write, snapshot or lint pass: the block does not pair its own first and last keywords"
    d("blockProcedureHeaders") = nh
    d("blockProcedureEnds") = ne
    d("extraEndCount") = extra
    d("extraEndKeywords") = "the block opens " & CStr(nh) & " procedure(s) and closes " & CStr(ne) & ", so " & CStr(extra) & " trailing End <Sub|Function|Property> line(s) have no matching header. VB6 reports exactly this as VB6-E105 (extra end function) and refuses to compile."
    d("suggestedAction") = "resubmit the SAME block without the " & CStr(extra) & " extra trailing terminator line(s) - the corrected text is in suggestedText - or split the work so each call carries one header and its own End."
    d("suggestedText") = fixed
    PreWritePairingRefusal = JsonStringify(d)
End Function

' The target procedure's first/last line, in PANE coordinates, from a pane TEXT
' (used before and after the edit so a line drift is visible in the receipt).
Private Function ProcEchoDict(ByVal text As String, ByVal line As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("proc") = ""
    d("available") = False
    d("startLine") = 0
    d("endLine") = 0
    d("firstLine") = ""
    d("lastLine") = ""
    If Len(text) = 0 Then
        Set ProcEchoDict = d
        Exit Function
    End If
    Dim names() As String, starts() As Long
    Vb6Lint.LintProcMap text, names, starts
    If line < 1 Or line > UBound(names) + 1 Then
        Set ProcEchoDict = d
        Exit Function
    End If
    Dim pn As String, ps As Long
    pn = names(line - 1)
    ps = starts(line - 1)
    d("proc") = pn
    d("startLine") = ps
    If Len(pn) = 0 Or ps <= 0 Then
        Set ProcEchoDict = d
        Exit Function
    End If
    Dim lines() As String
    Dim i As Long, pe As Long
    lines = TextLines(text)
    pe = ps
    For i = ps To UBound(lines) + 1
        If i > ps Then
            If Len(Vb6Lint.LintProcHeaderName(lines(i - 1))) > 0 Then Exit For
        End If
        pe = i
    Next i
    Do While pe > ps
        If Len(Trim$(lines(pe - 1))) = 0 Then
            pe = pe - 1
        Else
            Exit Do
        End If
    Loop
    d("available") = True
    d("endLine") = pe
    d("firstLine") = lines(ps - 1)
    d("lastLine") = lines(pe - 1)
    Set ProcEchoDict = d
End Function

' TASK-P1-126 (MCP-07): the refusal for a line-number edit with no expected_sha256.
' It carries the sha the caller needs, so the fix is ONE retry with the value in hand.
Private Function MissingShaRefusal(ByVal component As String, ByVal operation As String, _
                                   ByVal startLine As Long, ByVal endLine As Long, ByVal cm As VBIDE.CodeModule) As String
    Dim cur As String
    cur = ""
    On Error Resume Next
    If Not cm Is Nothing Then
        If cm.CountOfLines > 0 Then cur = cm.Lines(1, cm.CountOfLines)
    End If
    Err.Clear
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "expected_sha256_required"
    d("tool") = "replace_code"
    d("component") = component
    d("operation") = operation
    Set d("requestedRange") = RangeDict(operation, startLine, endLine)
    d("currentSha256") = Sha256.Sha256HexTextUtf8(cur)
    d("shaSpace") = "IDE code-pane text (the same sha get_component_code reports as sha256)"
    d("lineSpace") = "ide"
    d("lineSpaceHint") = "requestedRange is IDE code-pane lines; get_component_code's numbered output is DISK lines (ideLine = diskLine - coords.headerOffset)"
    d("note") = "REFUSED BEFORE any write: a replace/delete addressed by line number must pin the text it was measured against. Stale line numbers are how a remembered range silently edits ANOTHER procedure (a real loss on this project)."
    d("suggestedAction") = "call get_component_code to re-read the component, then resubmit with expected_sha256 set to the sha256 in that receipt (or to currentSha256 here if it is unchanged). To edit by content instead of by line number - which cannot drift - use edit_by_anchor."
    MissingShaRefusal = JsonStringify(d)
End Function

' TASK-P1-126 (MCP-15): baseline_diff. Keys are harvested from the caller's
' baseline TEXT, so the caller can pass either the plain stableKey lines or a whole
' previous validate_code receipt - no JSON parser is needed and no key is guessed.
' TASK-P1-127 (R2): the P1-126 harvester was broken in two independent ways, both
' measured by the independent verifier (内部验证回执 Y4/Y5, c3_log.json T4):
'   (a) it only accepted an occurrence PRECEDED BY A COMMA, so the documented "bare
'       stableKey lines" input harvested 0 keys (baselineKeyCount=0);
'   (b) after consuming the prefix it left the cursor ON the value's opening quote,
'       the value loop exited on its first character, and EVERY key was added EMPTY.
'       A whole receipt therefore reported baselineKeyCount=2 together with
'       addedCount=2 AND removedCount=2 and unchangedCount=0 - each pre-existing
'       diagnostic counted twice, which is exactly the MCP-15 symptom this field
'       exists to remove.
' The harvester below takes BOTH forms, starts the value one character PAST the
' opening quote, unescapes the escapes JsonLib can emit (\" \\ \n \r \t \uXXXX)
' and de-duplicates.
Private Function ExtractStableKeys(ByVal text As String) As Object
    Dim coll As Object
    Dim seen As Object
    Dim lines() As String
    Dim n As Long, i As Long, pos As Long, hit As Long
    Dim tok As String, k As String, t As String, c1 As String
    Dim ch As String, nx As String, hx As String
    Dim uCode As Long
    Set coll = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    n = Len(text)
    If n > 0 Then
        ' ---- (1) the JSON form: "stableKey":"<value>" ---------------------------
        tok = Chr$(34) & "stableKey" & Chr$(34) & ":" & Chr$(34)
        pos = 1
        Do
            hit = InStr(pos, text, tok, vbBinaryCompare)
            If hit = 0 Then Exit Do
            pos = hit + Len(tok)
            k = ""
            Do While pos <= n
                ch = Mid$(text, pos, 1)
                If ch = Chr$(34) Then
                    pos = pos + 1
                    Exit Do
                ElseIf ch = Chr$(92) And pos < n Then
                    nx = Mid$(text, pos + 1, 1)
                    Select Case nx
                        Case Chr$(34): k = k & Chr$(34): pos = pos + 2
                        Case Chr$(92): k = k & Chr$(92): pos = pos + 2
                        Case "n": k = k & Chr$(10): pos = pos + 2
                        Case "r": k = k & Chr$(13): pos = pos + 2
                        Case "t": k = k & Chr$(9): pos = pos + 2
                        Case "u"
                            hx = Mid$(text, pos + 2, 4)
                            uCode = -1
                            If Len(hx) = 4 Then
                                On Error Resume Next
                                Err.Clear
                                uCode = CLng("&H" & hx)
                                If Err.Number <> 0 Then uCode = -1
                                Err.Clear
                                On Error GoTo 0
                            End If
                            If uCode >= 0 Then
                                k = k & ChrW$(uCode)
                                pos = pos + 6
                            Else
                                k = k & nx
                                pos = pos + 2
                            End If
                        Case Else: k = k & nx: pos = pos + 2
                    End Select
                Else
                    k = k & ch
                    pos = pos + 1
                End If
            Loop
            If Len(k) > 0 Then
                If Not seen.Exists(k) Then
                    seen(k) = True
                    coll.Add k
                End If
            End If
        Loop
    End If
    ' ---- (2) bare stableKey lines: <ruleId>|<message>|<proc> --------------------
    lines = TextLines(text)
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Len(t) > 0 And InStr(t, "|") > 0 And InStr(t, Chr$(34)) = 0 Then
            c1 = Left$(t, 1)
            If c1 <> "{" And c1 <> "[" And c1 <> "}" And c1 <> "]" Then
                If Not seen.Exists(t) Then
                    seen(t) = True
                    coll.Add t
                End If
            End If
        End If
    Next i
    ' TASK-P1-126: Set is REQUIRED for an object-returning Function (error 91 without it).
    Set ExtractStableKeys = coll
End Function

Private Function BaselineDiffDict(ByVal baselineText As String, ByVal arr As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim keys As Object
    Set keys = ExtractStableKeys(baselineText)
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To keys.Count
        seen(CStr(keys(i))) = True
    Next i
    Dim added As Object, removed As Object
    Set added = New Collection
    Set removed = New Collection
    Dim unchanged As Long
    unchanged = 0
    Dim cur As Object
    Set cur = CreateObject("Scripting.Dictionary")
    If Not arr Is Nothing Then
        For i = 1 To arr.Count
            Dim it As Object
            Set it = arr(i)
            Dim k As String
            k = ""
            If Not it Is Nothing Then k = DictStr(it, "stableKey")
            If Len(k) > 0 Then
                cur(k) = True
                If seen.Exists(k) Then
                    unchanged = unchanged + 1
                Else
                    added.Add k
                End If
            End If
        Next i
    End If
    For i = 1 To keys.Count
        Dim bk As String
        bk = CStr(keys(i))
        If Not cur.Exists(bk) Then removed.Add bk
    Next i
    d("enabled") = True
    d("matchedOn") = "stableKey (ruleId|message|proc) - deliberately EXCLUDES procOffset, because inserting lines ABOVE a procedure shifts every line number while the diagnostic itself is unchanged"
    d("fingerprintField") = "fingerprint = ruleId|message|proc|procOffset (exact position inside the procedure); stableKey = the same without procOffset"
    d("baselineKeyCount") = keys.Count
    Set d("added") = added
    Set d("removed") = removed
    d("addedCount") = added.Count
    d("removedCount") = removed.Count
    d("unchangedCount") = unchanged
    d("note") = "added[] are diagnostics that are NOT in the baseline; a baseline diagnostic that only MOVED is not added (that was MCP-15: inserting 74 lines made 11 pre-existing false positives look new)."
    Set BaselineDiffDict = d
End Function

' The first and last line that differ between two texts (1-based, in the NEW text).
Private Function ChangedRegion(ByVal before As String, ByVal after As String) As Object
    Dim a() As String, b() As String
    a = TextLines(before)
    b = TextLines(after)
    Dim na As Long, nb As Long, pre As Long, suf As Long
    na = UBound(a) + 1
    nb = UBound(b) + 1
    pre = 0
    Do While pre < na And pre < nb
        MaybePump
        If a(pre) <> b(pre) Then Exit Do
        pre = pre + 1
    Loop
    suf = 0
    Do While suf < (na - pre) And suf < (nb - pre)
        MaybePump
        If a(na - 1 - suf) <> b(nb - 1 - suf) Then Exit Do
        suf = suf + 1
    Loop
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("start_line") = pre + 1
    d("end_line") = nb - suf
    If nb - suf < pre + 1 Then d("end_line") = pre
    Set ChangedRegion = d
End Function

' The text a replace_code operation is expected to produce, built from the pane
' BEFORE the IDE is touched, so beta-3 can lint the result (and alpha-3 can then
' compare the prediction with what the IDE really holds).
' TASK-P1-049: this used to rebuild the module with one "out = out & ..." per
' line (and one per line of the NEW text). On a 1.2 MB / 28.8k-line module that
' is quadratic: P4-039 measured 62,402 ms / 86,800 ms for a single 5->14 line
' replace, all of it here. Every piece is now produced in ONE pass with Join
' (JoinRange / JoinLines); the result is byte-identical, because the old loops
' appended a CRLF after every prefix line and between suffix lines, which is
' exactly what the Joins below produce.
Private Function PredictedReplaceText(ByVal curPane As String, ByVal op As String, ByVal startLine As Long, ByVal endLine As Long, ByRef newLines() As String, ByVal newCount As Long) As String
    Dim cur() As String
    cur = TextLines(curPane)
    Dim n As Long
    n = UBound(cur) + 1
    Dim body As String, out As String
    body = JoinLines(newLines, newCount)
    Select Case op
        Case "set"
            PredictedReplaceText = body
            Exit Function
        Case "append"
            out = curPane
            If Len(out) > 0 And newCount > 0 Then out = out & vbCrLf
            PredictedReplaceText = out & body
            Exit Function
        Case "replace"
            Dim rs As Long, re As Long
            rs = startLine
            re = endLine
            If rs < 1 Then rs = 1
            If re < rs Then re = rs
            If re > n Then re = n
            out = ""
            If rs > 1 Then out = JoinRange(cur, 0, rs - 2) & vbCrLf
            out = out & body
            If re < n Then out = out & vbCrLf & JoinRange(cur, re, n - 1)
            PredictedReplaceText = out
            Exit Function
        Case "insert"
            Dim ip As Long
            ip = startLine
            If ip < 1 Then ip = 1
            If ip > n + 1 Then ip = n + 1
            out = ""
            If ip > 1 Then out = JoinRange(cur, 0, ip - 2) & vbCrLf
            out = out & body
            If ip <= n Then out = out & vbCrLf & JoinRange(cur, ip - 1, n - 1)
            PredictedReplaceText = out
            Exit Function
        Case "delete"
            Dim ds As Long, de As Long
            ds = startLine
            de = endLine
            If ds < 1 Then ds = 1
            If de < ds Then de = ds
            If de > n Then de = n
            out = ""
            If ds > 1 Then out = JoinRange(cur, 0, ds - 2) & vbCrLf
            If de < n Then out = out & JoinRange(cur, de, n - 1)
            PredictedReplaceText = out
            Exit Function
    End Select
End Function

' TASK-P1-049: Join a BOUNDED slice of a line array in ONE pass. This is the
' replacement for the "out = out & ..." per-line loops, which are quadratic on a
' whole module (28.8k lines -> tens of seconds).
Private Function JoinRange(ByRef arr() As String, ByVal i0 As Long, ByVal i1 As Long) As String
    If i1 < i0 Then
        JoinRange = ""
        Exit Function
    End If
    Dim parts() As String, i As Long, k As Long
    ReDim parts(0 To i1 - i0)
    k = 0
    For i = i0 To i1
        MaybePump
        parts(k) = arr(i)
        k = k + 1
    Next
    JoinRange = Join(parts, vbCrLf)
End Function

' Snapshot of a build artifact so a receipt can say whether it was refreshed.
Private Sub ArtifactSnap(ByVal path As String, ByRef exists As Boolean, ByRef size As Double, ByRef mtime As Double)
    On Error Resume Next
    exists = False
    size = 0
    mtime = 0
    If Len(path) = 0 Then Exit Sub
    If Dir$(path) = "" Then Exit Sub
    size = FileLen(path)
    mtime = FileDateTime(path)
    exists = True
End Sub

Private Function ArtifactMtimeText(ByVal exists As Boolean, ByVal mtime As Double) As String
    If Not exists Then Exit Function
    ArtifactMtimeText = Format$(mtime, "yyyy-mm-dd hh:nn:ss")
End Function

Private Function ToolValidateCode(ByVal name As String, ByVal code As String, ByVal kind As String, ByVal mode As String, Optional ByVal baseline As String = "") As String
    If Len(mode) = 0 Then mode = "warn"
    If Len(kind) = 0 Then kind = InferKind(name)
    If LCase$(mode) = "off" Then
        ToolValidateCode = "{" & Chr$(34) & "ok" & Chr$(34) & ":true," & Chr$(34) & "mode" & Chr$(34) & ":" & Chr$(34) & "off" & Chr$(34) & "}"
        Exit Function
    End If
    ' TASK-P1-111 (S1): validate_code answered ok:false with NO structured reason - the last
    ' sibling of TASK-P1-110's check_component fix (LintResultJson gained wantRefusalReason for
    ' exactly this). ADDITIVE and FAILURE-ONLY: reason/suggestedAction/refusalDetail appear only
    ' when severity=error diagnostics exist; a clean module's receipt is byte-identical.
    ToolValidateCode = LintResultJson(name, kind, mode, code, True, baseline)
End Function

Private Function ToolCheckComponent(ByVal name As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & name
    Dim text As String
    Dim kind As String
    ' TASK-P1-062: guard the empty-component case (CountOfLines = 0 => Lines(1,0) = error 5)
    If cm.CountOfLines > 0 Then text = cm.Lines(1, cm.CountOfLines)
    If comp.Type = 1 Then
        kind = "bas"
    ElseIf comp.Type = 2 Then
        kind = "cls"
    Else
        kind = "frm"
    End If
    ToolCheckComponent = LintResultJson(comp.Name, kind, "warn", text, True)
End Function

Private Function ToolReadSource(ByVal path As String) As String
    Dim full As String, b() As Byte, enc As String, text As String
    full = ResolvePath(path)
    Encoding.ReadAllBytesTo full, b
    If Encoding.ByteCount(b) <= 0 Then Err.Raise vbObjectError + 70, "CodeTools", "File empty or unreadable: " & full
    enc = Encoding.DetectEncoding(b)
    Encoding.DecodeBytes b, enc, text
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("file") = full
    root("encoding") = enc
    root("bytes") = Encoding.ByteCount(b)
    root("sha256") = Sha256.Sha256HexBytes(b)
    Dim lines() As String
    lines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    root("lineCount") = UBound(lines) + 1
    root("text") = text
    ToolReadSource = JsonStringify(root)
End Function

' =====================================================================
' integration tools batch 2 (compile-verified)
' =====================================================================

Public Function JsonGetAny(ByVal args As Object, ByVal key As String) As Object
    On Error Resume Next
    If args Is Nothing Then Exit Function
    If args.Exists(key) Then
        Set JsonGetAny = args(key)
    End If
End Function

Private Function LineCount(ByVal s As String) As Long
    If Len(s) = 0 Then Exit Function
    LineCount = UBound(Split(Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf), vbLf)) + 1
End Function

Private Function CompText(ByVal name As String) As String
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Err.Raise vbObjectError + 51, "CodeTools", "Component has no code module: " & name
    ' TASK-P1-062: this is the CENTRE of the defect (20+ tools reach it). An empty module
    ' (CountOfLines = 0 - e.g. the default-named module a name-colliding import creates)
    ' made Lines(1, 0) raise VB error 5. Return "" instead; the output for NON-empty
    ' components is unchanged.
    If cm.CountOfLines > 0 Then CompText = cm.Lines(1, cm.CountOfLines)
End Function

Private Sub SetCompText(ByVal name As String, ByVal text As String)
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim cm As VBIDE.CodeModule
    Set cm = comp.CodeModule
    If cm Is Nothing Then Exit Sub
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    If Len(text) > 0 Then cm.InsertLines 1, text
    ' TASK-P1-021: clearing a code module leaves one physical empty line behind, so
    ' the pane ended up ONE line longer than the text that was written - the exact
    ' mechanism behind "the receipt lies about the line count" (measured in
    ' TASK-P1-020: newTotalLines 24 vs ideTotalLines 25 + syncWarning). Trim only
    ' trailing EMPTY lines the text does not itself have, so the IDE line count and
    ' the text line count agree by construction.
    Dim want As Long, guard As Long
    want = LineCount(text)
    guard = 0
    Do While cm.CountOfLines > want And guard < 8
        If Len(Trim$(cm.Lines(cm.CountOfLines, 1))) <> 0 Then Exit Do
        cm.DeleteLines cm.CountOfLines, 1
        guard = guard + 1
    Loop
End Sub

Private Sub GuardSha(ByVal cur As String, ByVal expected As String)
    If Len(expected) = 0 Then Exit Sub
    Dim actual As String
    actual = Sha256.Sha256HexTextUtf8(cur)
    If actual <> UCase$(expected) Then
        ' TASK-P1-027 (P4-026 finding): the old raise produced a BARE "Error: content
        ' changed: ..." with no ok/reason. Tag it with MCPREASON: so CallTool's Fail
        ' handler can answer with the same {ok:false, reason:...} shape as every other
        ' refusal - one change covers all nine guard call sites.
        Err.Raise vbObjectError + 81, "CodeTools", "MCPREASON:sha256_mismatch|" & UCase$(expected) & "|" & actual
    End If
End Sub

' TASK-P1-020 alpha: the edit receipt now states which line space every hash
' belongs to, what the IDE says the line count is (next to the text-derived one),
' what the linter says about the result, and which range was requested vs applied.
' 'pack' is the pre-write lint result (Nothing = lint the after text post-write);
' 'extra' carries requestedRange / appliedRange / clamped / hits when the caller
' has them.
Private Function ChangeResult2(ByVal name As String, ByVal before As String, ByVal after As String, ByVal pack As Object, ByVal extra As Object, Optional ByVal tEntry As Long = 0) As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = name
    root("sha256Before") = Sha256.Sha256HexTextUtf8(before)
    root("sha256After") = Sha256.Sha256HexTextUtf8(after)
    root("lineDelta") = LineCount(after) - LineCount(before)
    root("newTotalLines") = LineCount(after)
    Dim ideTotal As Long
    ideTotal = CompLineCount(name)
    root("ideTotalLines") = ideTotal
    If ideTotal <> LineCount(after) Then root("syncWarning") = SyncWarningText(ideTotal, LineCount(after))
    Set root("coords") = CoordsDictOfName(name, 1, "ide")
    ' TASK-P1-085: enumerate the project-directory files THIS call wrote (today: the _undo
    ' snapshot taken by the single text-write funnel). Always present, [] when nothing was
    ' written, so a caller never has to guess what the call left in the project directory.
    Dim pdwU As Object
    Set pdwU = New Collection
    ManifestMergeUndo pdwU
    Set root("projectDirWrites") = pdwU
    AddPaneDiskFields root, name, after
    If pack Is Nothing Then
        CopyLintPack root, LintPack(after, KindOfName(name))
    Else
        CopyLintPack root, pack
    End If
    If Not extra Is Nothing Then
        ' TASK-P1-024: merge EVERY extra field (matchedText/matchedLine/matchedLines,
        ' scope, ...) instead of a hand-written whitelist - a whitelist silently drops
        ' new fields, which is how matchedText/matchedLine went missing.
        Dim ek As Variant
        For Each ek In extra.Keys
            On Error Resume Next
            Err.Clear
            If IsObject(extra(ek)) Then
                Set root(ek) = extra(ek)
            Else
                root(ek) = extra(ek)
            End If
            On Error GoTo 0
        Next
    End If
    ' TASK-P1-057: totalMs covers the WHOLE call (entry -> receipt), so it must agree
    ' with the wall clock instead of hiding the receipt-building cost.
    If tEntry > 0 Then root("totalMs") = GetTickCount() - tEntry
    ChangeResult2 = ReceiptJson(root)
End Function

Private Function ChangeResult(ByVal name As String, ByVal before As String, ByVal after As String) As String
    ChangeResult = ChangeResult2(name, before, after, Nothing, Nothing)
End Function

' TASK-P1-086 (F3): serialise a receipt and stamp the two fields that make the
' receipt-building TAIL visible. The caller has already stamped totalMs (entry -> receipt
' dict ready), which keeps every historical totalMs number comparable; serializeMs and
' receiptBytes describe what happens AFTER that stamp. receiptBytes is the UTF-8 byte
' length of the receipt text the CLIENT will actually receive: JsonStringify produces
' the body, then the dispatch layer prepends the mcp/ide stamp (StampMcp) to every
' tools/call text payload, so the measured number is Utf8ByteLen(StampMcp(body)) - the
' same encoding and the same counter the HTTP writer uses for Content-Length
' (McpServer.BuildHttpResponse: Utf8Encode + UBound). That makes the field falsifiable
' from outside: len(content[0].text encoded as UTF-8) must equal receiptBytes.
' DEPENDENCY (stated, not hidden): the number is exact because the dispatcher stamps
' every tools/call text exactly once (CodeTools: item("text") = StampMcp(...)). If a
' future path ever returned a receipt WITHOUT that stamp, receiptBytes would overstate
' it by the stamp length. The JSON-RPC envelope, the socket, the hub hop and the client
' parse are all OUTSIDE this process: no field here claims to cover them, and the
' transport time is NOT measurable from inside the add-in.
' The loop re-serialises until the byte length stops moving, so the returned string
' really does carry its own receiptBytes. serializeMs is the ACCUMULATED duration of
' every convergence pass (serialise + stamp + byte count), not just the last one: that
' is the number that has to line up with "wall clock minus totalMs", which is how P1-074
' exposed the 7.3 s that no field covered. PRECISION, stated honestly: the pass that
' produces the returned string cannot report its own duration (the value would change
' the bytes it is embedded in), so the stamped total excludes exactly that last pass -
' i.e. serializeMs is the receipt tail minus one pass, and receiptBytes stays exact.
Private Function ReceiptJson(ByVal root As Object) As String
    Dim s As String
    Dim n As Long, guard As Long, tSer As Long, totalSer As Long
    root("serializeMs") = 0
    root("receiptBytes") = 0
    Do
        tSer = GetTickCount()
        s = JsonStringify(root)
        n = Utf8ByteLen(StampMcp(s))
        totalSer = totalSer + (GetTickCount() - tSer)
        ' The numbers already written into the string are final when re-encoding root
        ' with them yields the same byte length.
        If n = CLng(root("receiptBytes")) Then Exit Do
        root("serializeMs") = totalSer
        root("receiptBytes") = n
        guard = guard + 1
    Loop While guard < 6
    ReceiptJson = s
End Function

' TASK-P1-086 (F3): UTF-8 byte length of a string, using exactly the kernel32 conversion
' the transport uses, so receiptBytes and the eventual Content-Length agree.
Private Function Utf8ByteLen(ByVal s As String) As Long
    Utf8ByteLen = 0
    If Len(s) = 0 Then Exit Function
    Dim b() As Byte
    b = Utf8Encode(s)
    On Error Resume Next
    Utf8ByteLen = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Then Utf8ByteLen = 0
    On Error GoTo 0
End Function

Private Function CollectionCount(ByVal c As Object) As Long
    On Error Resume Next
    CollectionCount = 0
    CollectionCount = c.Count
End Function

Private Function DictStr(ByVal d As Object, ByVal k As String) As String
    On Error Resume Next
    DictStr = ""
    If d.Exists(k) Then DictStr = CStr(d(k))
End Function

' TASK-P1-082: the anchor-batch op record is TAB separated, and Chr$(1) already stands
' for CRLF inside a payload. A literal TAB in the caller's text used to SHIFT the
' fields: apply_edits then read the tail of the text as the occurrence field and died
' with VB error 13 ("类型不匹配"), while edit_by_anchor / preview_edit truncated the
' text at the TAB without saying anything. Escape the separator on the way in and
' restore it on the way out. With no TAB in the input every Replace below is a no-op,
' so existing calls keep byte-identical behaviour. (A raw Chr$(1)/Chr$(2) could never be
' represented through this record either - that is the pre-existing convention, not a
' new restriction: VB6 source text cannot contain those control characters.)
Private Function OpEsc(ByVal s As String) As String
    OpEsc = Replace(s, vbTab, Chr$(2))
End Function

Private Function OpUnesc(ByVal s As String) As String
    OpUnesc = Replace(s, Chr$(2), vbTab)
End Function

Private Function BuildOps(ByVal edits As Object) As String()
    Dim n As Long, ops() As String, i As Long
    n = CollectionCount(edits)
    ' TASK-P1-100: "no edits" is CALLER INPUT, not an add-in fault. It used to be a bare raise
    ' that CallToolInner could only wrap as {ok:false, reason:"unexpected_error"}; tag it so the
    ' MCPREASON channel answers with a branchable refusal. BuildOps has exactly TWO callers
    ' (apply_edits and preview_edit), so no other tool's receipt can move.
    If n = 0 Then Err.Raise vbObjectError + 83, "CodeTools", "MCPREASON:edits_required"
    ReDim ops(0 To n - 1)
    i = 0
    Dim it As Variant, d As Object, op As String, anc As String, txt As String
    Dim occ As Long, mch As String
    For Each it In edits
        Set d = it
        op = UCase$(DictStr(d, "op"))
        anc = DictStr(d, "anchor")
        txt = Replace(DictStr(d, "text"), vbCrLf, Chr$(1))
        occ = JsonGetLong(d, "occurrence", 0)
        mch = DictStr(d, "match")
        ' TASK-P1-024 gamma-5/6: occurrence (1-based) and match ride along in the op
        ' entry so the resolver can pick the right hit for a repeated line.
        ' TASK-P1-082: every free-text field is escaped so a TAB in it cannot split the
        ' record (occurrence stays the 4th field, so CLng below can never see text).
        ops(i) = op & vbTab & OpEsc(anc) & vbTab & OpEsc(txt) & vbTab & CStr(occ) & vbTab & OpEsc(mch)
        i = i + 1
    Next
    BuildOps = ops
End Function

Private Function KindOfComponent(ByVal comp As VBIDE.VBComponent) As String
    If comp.Type = 1 Then
        KindOfComponent = "bas"
    ElseIf comp.Type = 2 Then
        KindOfComponent = "cls"
    Else
        KindOfComponent = "frm"
    End If
End Function

Private Function BaseNameOf(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, Chr$(92))
    If p = 0 Then
        BaseNameOf = path
    Else
        BaseNameOf = Mid$(path, p + 1)
    End If
End Function

Private Function HeaderOffsetForPath(ByVal path As String, ByVal ideLines As Long) As Long
    Dim b() As Byte, enc As String, text As String, dl As Long
    Encoding.ReadAllBytesTo path, b
    If Encoding.ByteCount(b) <= 0 Then Exit Function
    enc = Encoding.DetectEncoding(b)
    Encoding.DecodeBytes b, enc, text
    dl = LineCount(text)
    If Len(text) > 0 Then
        If Right$(text, 1) = vbLf Or Right$(text, 1) = vbCr Then dl = dl - 1
    End If
    HeaderOffsetForPath = dl - ideLines
End Function

Private Function DiskPathOf(ByVal name As String, ByVal given As String) As String
    If Len(given) > 0 Then
        DiskPathOf = ResolvePath(given)
        Exit Function
    End If
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    On Error Resume Next
    DiskPathOf = comp.FileNames(1)
    On Error GoTo 0
End Function

Private Function HeaderOffsetForDiskName(ByVal file As String) As Long
    Dim bn As String
    HeaderOffsetForDiskName = -1
    If Len(file) = 0 Then Exit Function
    bn = LCase$(BaseNameOf(file))
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim c As VBIDE.VBComponent, fn As String
    For Each c In p.VBComponents
        fn = ""
        On Error Resume Next
        fn = c.FileNames(1)
        On Error GoTo 0
        If Len(fn) > 0 Then
            If LCase$(BaseNameOf(fn)) = bn Then
                HeaderOffsetForDiskName = HeaderOffsetForPath(fn, c.CodeModule.CountOfLines)
                Exit Function
            End If
        End If
    Next
End Function

Private Function ToolBuild(ByVal project As String, ByVal defines As String, ByVal outPath As String, ByVal timeoutMs As Long, _
                          ByVal laaGiven As Boolean, ByVal laaParam As Boolean, ByVal diagnoseWant As Boolean) As String
    Dim vbp As String, logPath As String, cmdArgs As String
    ' Card TASK-P1-015: the caller is blocked for the WHOLE call (the MCP server is
    ' single threaded), so start counting at entry. prepMs then means everything
    ' before the compiler starts, compileMs the compiler subprocess itself.
    Dim el0 As Long
    el0 = GetTickCount()
    If Len(project) > 0 Then
        vbp = ResolvePath(project)
    Else
        vbp = GetProject().FileName
    End If
    If Len(vbp) = 0 Then Err.Raise vbObjectError + 86, "CodeTools", "project has no file name; save the project first"
    ' Card TASK-P1-014: fail fast. Never hand a non-existent .vbp to VB6.EXE /make -
    ' that is what turned a mistyped path into a hung build with a saturated CPU and
    ' no readable error. Answered structurally, without launching anything.
    ' Dir$ raises VB error 52 ("bad file name or number") for some syntactically odd
    ' targets (measured: a UNC path). Treat that as "not found" too, so every invalid
    ' input still yields the structured answer instead of a raw VB error.
    Dim notFound As Boolean
    notFound = False
    On Error Resume Next
    Err.Clear
    notFound = (Dir$(vbp) = "")
    If Err.Number <> 0 Then notFound = True: Err.Clear
    On Error GoTo 0
    If notFound Then
        Dim nf As Object
        Set nf = CreateObject("Scripting.Dictionary")
        nf("ok") = False
        nf("reason") = "project_not_found"
        nf("path") = vbp
        nf("requested") = project
        nf("note") = "the resolved project path does not exist; VB6.EXE was NOT launched"
        ToolBuild = JsonStringify(nf)
        Exit Function
    End If
    ' ---- TASK-P1-132 (A10 sidecar): resolve WHERE the LAA decision comes from ----
    Dim laaWant As Boolean, laaSrc As String, laaProjFile As String, laaSrcNote As String, laaCfgErr As String
    Dim laaUsable As Boolean
    laaUsable = LaaResolve(vbp, laaGiven, laaParam, laaWant, laaSrc, laaProjFile, laaSrcNote, laaCfgErr)
    If timeoutMs <= 0 Then timeoutMs = 180000
    ' Card TASK-P1-008 (criterion 3): when the requested project IS the one open in
    ' the IDE, an isolated VB6.EXE /make cannot compile it - the IDE holds the file
    ' (measured: exit 1 with no error text). Prefer the in-IDE Make path there and
    ' report which route was taken.
    ' Card TASK-P1-015: collapse the project to ONE identity before anything else, so
    ' the IDE match, the /make command line, the artifact path and the cache key all
    ' speak about the same file however the caller spelled it (junction alias,
    ' relative form, separators, case). Identity, not spelling.
    Dim vbpCanon As String
    vbpCanon = Build.CanonicalPath(vbp)
    If Len(vbpCanon) > 0 Then vbp = vbpCanon
    Dim useIde As Boolean
    useIde = False
    Dim activeFile As String
    activeFile = ActiveProjectFile()
    If Len(activeFile) > 0 Then
        ' The IDE reports the project under whatever spelling opened it (measured:
        ' the IDE held the junction alias while the caller named the target). A
        ' spelling comparison missed that and routed a held project to an isolated
        ' /make, which cannot compile a file the IDE holds.
        If PathsSame(activeFile, vbp) Then useIde = True
    End If
    If Len(outPath) > 0 Then
        logPath = ResolvePath(outPath)
    Else
        logPath = Environ$("TEMP") & Chr$(92) & "vb6mcp_build.log"
    End If
    ' NOTE: the compiler log is NOT deleted here. The cache short-circuit below
    ' returns without compiling and hands back the previous logPath, so deleting
    ' it up front would destroy the artifact that response points at.
    cmdArgs = "/make " & Chr$(34) & vbp & Chr$(34) & " /out " & Chr$(34) & logPath & Chr$(34)
    Dim d() As String, i As Long
    ' TASK-P1-162 (R7-4): the /d arguments are recorded STRUCTURALLY as they are appended,
    ' so the receipt can show the exact command line and the exact macro list instead of
    ' asking the caller to trust that the define arrived.
    Dim ccArgs As Object
    Set ccArgs = New Collection
    Dim ccExtra As String
    ccExtra = ""
    If Len(defines) > 0 Then
        d = Split(defines, ",")
        For i = 0 To UBound(d)
            If Len(Trim$(d(i))) > 0 Then
                ' TASK-P1-162 (R7-4 / MCP-lead-4): do NOT wrap the /d value in quotes. VB6's
                ' compiler folds the quoted text into the macro name, so the macro silently
                ' never applies while the build still reports success. Quote ONLY when the
                ' value really contains a space (then there is something to protect).
                Dim dv As String
                dv = Trim$(d(i))
                Dim ccQuoted As Boolean
                Dim ccFrag As String
                ccQuoted = (InStr(1, dv, " ") > 0)
                If ccQuoted Then
                    ccFrag = " /d " & Chr$(34) & dv & Chr$(34)
                Else
                    ccFrag = " /d " & dv
                End If
                cmdArgs = cmdArgs & ccFrag
                ccExtra = ccExtra & ccFrag
                Dim ccOne As Object
                Set ccOne = CreateObject("Scripting.Dictionary")
                ccOne("raw") = dv
                Dim ccEq As Long
                ccEq = InStr(dv, "=")
                If ccEq > 0 Then
                    ccOne("name") = Left$(dv, ccEq - 1)
                    ccOne("value") = Mid$(dv, ccEq + 1)
                Else
                    ccOne("name") = dv
                    ccOne("value") = ""
                End If
                ccOne("quoted") = ccQuoted
                ccOne("argument") = Trim$(ccFrag)
                ccArgs.Add ccOne
            End If
        Next
    End If
    ' ---- Card TASK-P1-008: fingerprint, cache, changed ------------------------
    Dim comps As Object
    Dim fp As Object
    Dim cache As Object
    Dim changed As Object
    Dim buildAdvised As Boolean
    Dim advisedWhy As String
    Dim haveCache As Boolean
    Dim prepErr As String
    prepErr = ""
    buildAdvised = True
    advisedWhy = "no previous successful build for this project"
    haveCache = False
    Set changed = New Collection
    ' The incremental feedback is an OPTIMISATION: it must never be able to break
    ' a build. Every step is best effort and any failure simply falls through to a
    ' normal full build (and is surfaced in prepError for diagnosis).
    On Error Resume Next
    Set fp = Build.BuildFingerprint(vbp, defines, comps)
    If Err.Number <> 0 Then prepErr = "fingerprint: " & Err.Description: Err.Clear
    On Error GoTo 0
    Set cache = CreateObject("Scripting.Dictionary")
    cache("key") = "": cache("hash") = "": cache("ok") = False: cache("exitCode") = 0
    cache("artifact") = "": cache("size") = 0: cache("logPath") = "": cache("at") = 0
    cache("verdict") = "": cache("okByKeyword") = False
    On Error Resume Next
    haveCache = Build.CacheLoad(cache)
    If Err.Number <> 0 Then prepErr = prepErr & " cacheLoad: " & Err.Description: Err.Clear
    If Not fp Is Nothing Then
        ' Only diff against a record for THIS project: the cache holds a single
        ' last-build record, and comparing another project's component list would
        ' report every component as added/removed.
        If haveCache Then
            If CStr(cache("key")) = CStr(fp("key")) Then
                Set changed = Build.ChangedComponents(comps, cache)
                If Err.Number <> 0 Then prepErr = prepErr & " changed: " & Err.Description: Err.Clear
            End If
        End If
    End If
    On Error GoTo 0
    If (Not fp Is Nothing) And haveCache Then
        If CStr(cache("key")) = CStr(fp("key")) Then
            If CStr(cache("hash")) = CStr(fp("hash")) And CBool(cache("ok")) Then
                buildAdvised = False
                advisedWhy = "project fingerprint matches the last successful build"
            ElseIf changed.Count > 0 Then
                advisedWhy = CStr(changed.Count) & " component(s) differ from the last build"
            Else
                advisedWhy = "project inputs changed (defines or .vbp)"
            End If
        Else
            advisedWhy = "no previous build recorded for this project"
        End If
    End If
    ' Short-circuit: identical inputs AND the artifact the last build produced is
    ' still on disk at the same size. VB6.EXE is not launched at all.
    If Not buildAdvised Then
        Dim szC As Double
        szC = 0
        If Build.ArtifactBuilt(CStr(cache("artifact")), 0, szC) Then
            If szC = CDbl(cache("size")) Then
                Dim croot As Object
                Set croot = CreateObject("Scripting.Dictionary")
                croot("ok") = True
                croot("cached") = True
                croot("exitCode") = CLng(cache("exitCode"))
                croot("file") = ""
                croot("reportedLine") = 0
                croot("diskLine") = 0
                croot("key") = ""
                croot("message") = ""
                croot("logPath") = CStr(cache("logPath"))
                croot("artifact") = CStr(cache("artifact"))
                croot("artifactSize") = szC
                ' TASK-P1-020 alpha: a cache hit did not compile anything, so the
                ' artifact is by definition unchanged - say so instead of implying
                ' a fresh build produced it.
                croot("artifactPath") = CStr(cache("artifact"))
                croot("artifactChanged") = False
                croot("artifactUnchangedReason") = "cache hit: VB6.EXE was not launched, so nothing rewrote the artifact"
                Dim cArtEx As Boolean, cArtSz As Double, cArtTm As Double
                ArtifactSnap CStr(cache("artifact")), cArtEx, cArtSz, cArtTm
                If cArtEx Then croot("artifactMtime") = ArtifactMtimeText(True, cArtTm)
                croot("okByKeyword") = CBool(cache("okByKeyword"))
                croot("verdict") = CStr(cache("verdict"))
                croot("path") = "cache"
                croot("elapsedMs") = GetTickCount() - el0
                Set croot("changed") = Build.NamesOf(changed)
                croot("buildAdvised") = False
                croot("advisedReason") = advisedWhy
                croot("projectPath") = vbp
                ' TASK-P1-125: the cached path publishes the SAME key set as a real
                ' build, so a caller never has to branch on which route answered.
                Dim cLog As String
                cLog = ""
                On Error Resume Next
                If Len(CStr(cache("logPath"))) > 0 Then
                    If Dir$(CStr(cache("logPath"))) <> "" Then cLog = Build.ReadLogAnsi(CStr(cache("logPath")))
                End If
                On Error GoTo 0
                croot("logAvailable") = (Len(cLog) > 0)
                croot("logBytes") = Len(cLog)
                If Len(cLog) > 60000 Then
                    croot("log") = Left$(cLog, 60000)
                    croot("logTruncated") = True
                Else
                    croot("log") = cLog
                    croot("logTruncated") = False
                End If
                If Len(cLog) = 0 Then croot("logUnavailableReason") = "no readable compiler log at the recorded logPath"
                croot("errorCount") = Build.CountBuildErrors(cLog)
                Dim cfe As Object
                Set cfe = CreateObject("Scripting.Dictionary")
                cfe("file") = ""
                cfe("line") = 0
                cfe("lineSpace") = "disk"
                cfe("text") = ""
                Set croot("firstError") = cfe
                Set croot("output") = ArtifactFullSnap(CStr(cache("artifact")))
                croot("artifact_locked") = False
                croot("artifact_locker_pid") = 0
                croot("artifact_locked_advice") = ""
                croot("pending") = False
                croot("artifact_holder_pid") = 0
                croot("route") = "cache"
                Dim cbk As Object
                Set cbk = CreateObject("Scripting.Dictionary")
                cbk("serverBlockedMs") = GetTickCount() - el0
                cbk("prepMs") = GetTickCount() - el0
                cbk("compileMs") = 0
                cbk("timeoutMs") = timeoutMs
                cbk("timeoutApplies") = False
                cbk("note") = "cache hit: no VB6.EXE subprocess was started, so timeout_ms (which bounds only a compiler subprocess) had nothing to bound"
                Set croot("blocking") = cbk
                croot("note") = "no component changed since the last successful build; VB6.EXE was not launched; timeout_ms bounds only the compiler subprocess and this route never starts one"
                ' TASK-P1-162 (R7-4): even a cache hit reports what WOULD have been passed.
                CcArgsInto croot, cmdArgs, ccExtra, ccArgs, defines, useIde
                ' TASK-P1-132 (A10): report (and, when needed, set) the LAA flag even on
                ' a cache hit, so laaBefore/laaAfter/laaSet exist on EVERY path.
                If laaUsable Then
                    ApplyLaaInto croot, CStr(cache("artifact")), laaWant
                    If CBool(croot("laaSet")) Then croot("laaNote") = "cache hit: no compile ran, but the artifact on disk was missing IMAGE_FILE_LARGE_ADDRESS_AWARE and its PE header was patched at rest"
                Else
                    LaaConfigRefusal croot, laaProjFile, laaCfgErr
                End If
                croot("laaSource") = laaSrc
                croot("laaProjectFile") = laaProjFile
                croot("laaSourceNote") = laaSrcNote
                croot("laaConfigError") = laaCfgErr
                ToolBuild = JsonStringify(croot)
                Exit Function
            End If
        End If
    End If
    ' Card TASK-P1-015: the caller is blocked for the WHOLE call, so report what the
    ' blocking time was spent on - preprocessing (fingerprint + cache) versus the
    ' compiler subprocess. elPrep marks the end of preprocessing.
    Dim elPrep As Long
    elPrep = GetTickCount()
    Dim ec As Long
    Dim tBefore As Double
    tBefore = Now
    ' Card TASK-P1-007 (i18n): decide from the artifact, not from a localized word
    Dim artifact As String
    artifact = Build.ArtifactPathOf(vbp)
    ' TASK-P1-020 alpha: a project without an ExeName32 line has no artifact path in
    ' the .vbp; on the in-IDE route the Make target IS derivable (the same rule
    ' make_project uses), so use it instead of reporting "unknown".
    If Len(artifact) = 0 And useIde Then
        On Error Resume Next
        artifact = MakeTargetPath(GetProject(), "")
        On Error GoTo 0
    End If
    ' TASK-P1-020 alpha: artifact BEFORE / AFTER, so a "successful" compile that did
    ' not refresh the artifact (locked, or identical bytes) is visible.
    Dim aBeforeExists As Boolean, aBeforeSize As Double, aBeforeTime As Double
    ArtifactSnap artifact, aBeforeExists, aBeforeSize, aBeforeTime
    Dim ideMakeJson As String
    ideMakeJson = ""
    Dim br As TBuildResult
    If useIde Then
        ' in-IDE Make: no compiler log exists on this route, so the freshly written
        ' artifact is the whole verdict
        br.ok = False
        br.okByKeyword = False
        br.verdict = "ide-make"
        br.raw = ""
        br.message = ""
        br.key = ""
        br.file = ""
        br.reportedLine = 0
        br.diskLine = 0
        br.exitCode = 0
        ideMakeJson = ToolMakeProject("", "ide", "", 0, False, laaGiven, laaParam, diagnoseWant)
        ec = 0
        Build.ApplyArtifactVerdict br, artifact, tBefore
    Else
        On Error Resume Next
        Kill logPath
        On Error GoTo 0
        If Not Build.RunExe(m_vbe.FullName, cmdArgs, timeoutMs, ec) Then Err.Raise vbObjectError + 87, "CodeTools", "build process failed to run"
        br = Build.ParseBuild(logPath, -1)
        Build.ApplyArtifactVerdict br, artifact, tBefore
    End If
    Dim aAfterExists As Boolean, aAfterSize As Double, aAfterTime As Double
    ArtifactSnap artifact, aAfterExists, aAfterSize, aAfterTime
    Dim artChanged As Boolean
    artChanged = False
    If aAfterExists Then
        If Not aBeforeExists Then
            artChanged = True
        ElseIf aAfterSize <> aBeforeSize Or aAfterTime <> aBeforeTime Then
            artChanged = True
        End If
    End If
    Dim off As Long
    off = HeaderOffsetForDiskName(br.file)
    ' TASK-P1-125: a project the IDE does not hold has no live header offset, so the
    ' disk line was 0. Count the hidden header from the file on disk instead.
    If off < 0 And Len(br.file) > 0 Then off = Build.HeaderOffsetFromFile(br.file)
    If off >= 0 And br.reportedLine > 0 Then br.diskLine = br.reportedLine + off + 1
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("ok") = br.ok
    root("exitCode") = ec
    root("file") = br.file
    root("reportedLine") = br.reportedLine
    root("diskLine") = br.diskLine
    ' ruling 4: this response carries line numbers whenever the compiler located an
    ' error, so it must carry coords too (reportedLine is 0-based IDE, diskLine 1-based)
    If br.reportedLine > 0 Then Set root("coords") = CoordsDict(br.reportedLine + 1, off, "ide", (off >= 0))
    root("key") = br.key
    root("message") = br.message
    If Not br.ok And Len(br.message) = 0 Then
        root("message") = "isolated /make reported no error text (exit code " & CStr(ec) & "); the compiler log is at logPath"
    End If
    root("logPath") = logPath
    ' Card TASK-P1-007 (i18n): expose how the verdict was reached
    root("artifact") = artifact
    ' Card TASK-P1-011: artifactSize must not present a STALE file's size as though a
    ' build had produced it. It is meaningful only when a fresh artifact was verified;
    ' otherwise it is 0 and any old size moves to staleArtifactSize.
    If CStr(br.verdict) = "artifact" Then
        root("artifactSize") = br.artifactSize
    Else
        root("artifactSize") = 0
        If br.artifactSize > 0 Then root("staleArtifactSize") = br.artifactSize
    End If
    root("okByKeyword") = br.okByKeyword
    root("verdict") = br.verdict
    ' ---- TASK-P1-020 alpha: compiled is NOT the same as artifact updated ----
    root("artifactPath") = artifact
    If Len(artifact) = 0 Then
        ' no artifact path: artifactChanged cannot be judged, and claiming false
        ' would read like the "locked" verdict - say it is unknown instead
        root("artifactChecked") = False
        root("artifactNote") = "the project declares no ExeName32 and no in-IDE Make target, so the artifact path is unknown and artifactChanged cannot be judged"
    Else
        root("artifactChanged") = artChanged
        If aAfterExists Then root("artifactMtime") = ArtifactMtimeText(True, aAfterTime)
    End If
    If Len(artifact) > 0 And Not artChanged Then
        ' The compiler can compile cleanly and still fail to WRITE the artifact when
        ' another process holds it. That case reports NO compile-error location
        ' (br.file stays empty), so it is distinguishable from a real compile error
        ' in a language-independent way. A located error is never called "compiled".
        Dim locatedErr As Boolean
        locatedErr = (Len(br.file) > 0)
        Dim writeOnlyFail As Boolean
        writeOnlyFail = (ec = 0) Or br.okByKeyword Or (Len(Trim$(br.message)) = 0)
        If aAfterExists And (Not locatedErr) And writeOnlyFail Then
            root("ok") = True
            root("verdict") = "compiled-no-artifact-change"
            root("artifactLockedHint") = "the code compiled but the artifact was NOT rewritten: it is most likely held open by another process (close it), or the build wrote identical bytes"
            ' TASK-P1-132 (MCP-lead-2 item 1): every exception above is a HYPOTHESIS.
            ' On this route the receipt cannot tell them apart from a compile failure,
            ' so say that instead of implying the build was merely redundant.
            root("unbuiltReasonUnknown") = True
            root("artifactNote") = "NOT built: the artifact was neither produced nor rewritten, and this route cannot say WHY - a compile failure is possible (compiled/okByKeyword are not evidence). Use make_project with diagnose=true, or build a COPY of the project with an isolated VB6.EXE /make, to get a real compiler log and firstError."
        End If
    End If
    ' ---- Card TASK-P1-008: incremental feedback -----------------------------
    ' ---- Card TASK-P1-008: incremental feedback -----------------------------
    root("cached") = False
    If useIde Then root("path") = "make_project" Else root("path") = "make"
    ' Card TASK-P1-015: state the route explicitly and account for the blocking time.
    ' TASK-P1-162 (R7-4): the actual command line + the structured /d list, so the
    ' "did my macro really get passed?" question is answerable from the receipt alone.
    CcArgsInto root, cmdArgs, ccExtra, ccArgs, defines, useIde
    root("projectPath") = vbp
    If useIde Then
        root("route") = "make_project"
    Else
        root("route") = "make"
    End If
    ' TASK-P1-127 (R5): make_project{route:"cli"} delegates HERE, and an isolated
    ' VB6.EXE /make cannot compile a project the IDE holds - so that call silently
    ' becomes an in-IDE Make. Say so, instead of letting the caller believe /make ran,
    ' and stop presenting the in-IDE route's exitCode 0 as a compiler verdict.
    Dim cliDelegated As Boolean
    cliDelegated = (LCase$(m_makeRouteOverride) = "ide")
    root("requestedRoute") = ""
    root("routeDowngraded") = False
    root("exitCodeApplies") = (Not useIde)
    If cliDelegated Then
        root("requestedRoute") = "cli"
        If useIde Then
            root("routeDowngraded") = True
            root("routeDowngradeReason") = "route=cli asked for an isolated VB6.EXE /make, but this project is open in the IDE and the IDE holds the file, so the isolated compiler cannot take it. The call fell back to the in-IDE Make (VBProject.MakeCompiledFile), which starts NO compiler subprocess and therefore writes no /out log."
        End If
    End If
    If useIde Then
        root("exitCodeNote") = "exitCode is 0 because NO compiler subprocess ran on this route (VBProject.MakeCompiledFile reports only success or failure). Do NOT read exitCode 0 as ""the compiler found no errors"" - read verdict, artifactChanged and output{} instead."
    End If
    Dim bk As Object
    Set bk = CreateObject("Scripting.Dictionary")
    bk("serverBlockedMs") = GetTickCount() - el0
    bk("prepMs") = elPrep - el0
    bk("compileMs") = GetTickCount() - elPrep
    If useIde Then
        bk("timeoutMs") = 0
        bk("timeoutApplies") = False
        bk("note") = "in-IDE Make: no compiler subprocess, so timeout_ms does not apply to this route"
    Else
        bk("timeoutMs") = timeoutMs
        bk("timeoutApplies") = True
        bk("note") = "timeout_ms bounds ONLY the compiler subprocess (VB6.EXE /make); the fingerprint/cache preprocessing is not bounded by it"
    End If
    Set root("blocking") = bk
    If Len(ideMakeJson) > 0 Then root("ideMake") = ideMakeJson
    root("elapsedMs") = GetTickCount() - el0
    If fp Is Nothing Then
        Set root("changed") = New Collection
    Else
        Set root("changed") = Build.NamesOf(changed)
    End If
    root("buildAdvised") = True
    root("advisedReason") = advisedWhy
    If Len(prepErr) > 0 Then root("prepError") = prepErr
    root("note") = "one build reports only the first error; timeout_ms bounds only the compiler subprocess (VB6.EXE /make), never the fingerprint/cache preprocessing and never an in-IDE Make"
    ' remember the successful fingerprint so the next identical call short-circuits
    If fp Is Nothing Then GoTo skipCache
    Dim ncache As Object
    Set ncache = CreateObject("Scripting.Dictionary")
    ncache("key") = fp("key")
    ncache("hash") = fp("hash")
    ncache("ok") = br.ok
    ncache("exitCode") = ec
    ncache("artifact") = artifact
    ncache("size") = br.artifactSize
    ncache("logPath") = logPath
    ncache("at") = Now
    ncache("verdict") = br.verdict
    ncache("okByKeyword") = br.okByKeyword
    Set ncache("comps") = comps
    Build.CacheSave ncache
skipCache:
    If Not br.ok Then
        ' Classic failure: the project is OPEN IN THE IDE, so the isolated
        ' VB6.EXE /make subprocess cannot take the file (the log says the file
        ' is already open) and message often ends up empty. Route the caller.
        root("routeHint") = "If this project is open in the VB6 IDE, use make_project instead: an isolated VB6.EXE /make cannot compile a file the IDE holds."
    End If
    ' ---- TASK-P1-125 (MCP-18 / MCP-28): the compiler log, the error count and the
    ' first error location. User evidence: the CLI route reported "compile error, line
    ' 26955 : sub or function not defined" while make_project answered error:"" - so the
    ' log TEXT, an error COUNT and the FIRST error location must be in the receipt, and
    ' logPath must be proven READABLE instead of merely echoed.
    Dim logText As String
    Dim logExists As Boolean
    Dim logSize As Double
    logText = ""
    logExists = False
    logSize = 0
    On Error Resume Next
    If Len(logPath) > 0 Then
        If Dir$(logPath) <> "" Then
            logExists = True
            logSize = CDbl(FileLen(logPath))
            logText = Build.ReadLogAnsi(logPath)
        End If
    End If
    On Error GoTo 0
    ' TASK-P1-127 (R5): the /out log is written ONLY by the isolated subprocess. On
    ' the in-IDE route no subprocess runs, so a file sitting at logPath is a LEFTOVER
    ' from an earlier build. P1-126 published it as THIS call's log (logAvailable=true,
    ' and errorCount=1 counted from the stale file) while firstError stayed empty and
    ' exitCode read 0 - exactly the "silent exit 0 + empty first error" F5 measured.
    ' The two are separated below, and the unreachable diagnosis is stated as such.
    Dim logFromThisCall As Boolean
    logFromThisCall = (Not useIde) And logExists
    root("logWrittenByThisCall") = logFromThisCall
    root("logAvailable") = logFromThisCall
    root("logBytes") = 0
    If logFromThisCall Then root("logBytes") = logSize
    If logExists And Not logFromThisCall Then
        root("staleLogPath") = logPath
        root("staleLogBytes") = logSize
        root("staleLogErrorCount") = Build.CountBuildErrors(logText)
        root("staleLogNote") = "a file exists at logPath, but NO compiler subprocess ran on this call, so those bytes are from an EARLIER build and are reported only as staleLog* - they are NOT this call's diagnosis."
    End If
    If logFromThisCall Then
        If Len(logText) > 60000 Then
            root("log") = Left$(logText, 60000)
            root("logTruncated") = True
            root("logTruncatedAt") = 60000
        Else
            root("log") = logText
            root("logTruncated") = False
        End If
    Else
        root("log") = ""
        If useIde Then
            root("logUnavailableReason") = "the in-IDE Make route starts NO compiler subprocess, so it writes no /out log; logPath is echoed for reference only"
        Else
            root("logUnavailableReason") = "no readable compiler log at logPath"
        End If
    End If
    root("errorCount") = 0
    If logFromThisCall Then root("errorCount") = Build.CountBuildErrors(logText)
    ' TASK-P1-127 (R5): reachable vs unreachable diagnosis, as a field rather than prose.
    root("diagnosisAvailable") = logFromThisCall
    root("diagnosisUnavailableReason") = ""
    root("diagnosisAlternative") = ""
    If Not logFromThisCall Then
        If useIde Then
            root("diagnosisUnavailableReason") = "this route (in-IDE Make / VBProject.MakeCompiledFile) has NO compiler subprocess, so no /out log and no compiler error text can exist for it - structurally, not as a failure of this call"
            root("diagnosisAlternative") = "run the isolated compiler where the IDE does NOT hold the project: call build{project=<a COPY of the .vbp>}, or make_project on a copy. An isolated VB6.EXE /make cannot compile a file the IDE has open - it exits 0 without writing the log."
        Else
            root("diagnosisUnavailableReason") = "the isolated /make subprocess ran but wrote no readable log at logPath"
            root("diagnosisAlternative") = "re-run and check that logPath is writable, or read firstError: when the compiler located an error it is still populated from the process output."
        End If
    End If
    Dim feOut As Object
    Set feOut = CreateObject("Scripting.Dictionary")
    feOut("file") = br.file
    feOut("line") = br.diskLine
    feOut("lineSpace") = "disk"
    feOut("lineIde") = br.reportedLine
    feOut("text") = br.message
    Set root("firstError") = feOut
    If Len(br.file) = 0 And Len(Trim$(br.message)) > 0 Then
        root("firstErrorNote") = "the compiler reported a message without a file location, so firstError.file is empty; firstError.text still carries the diagnostic"
    End If
    ' ---- TASK-P1-125 (3.2): is the artifact held by a running process? ----------
    Set root("output") = ArtifactFullSnap(artifact)
    root("pending") = (Not CBool(root("output")("exists"))) And br.ok And Len(artifact) > 0
    If CBool(root("pending")) Then root("pending_how_to_confirm") = "the artifact has not appeared yet: re-read output.exists / output.size / output.sha256 and check the file starts with a PE header (MZ at offset 0); do not read compiled/ok as evidence."
    root("artifact_locked") = False
    root("artifact_locker_pid") = 0
    root("artifact_locked_advice") = ""
    root("artifact_holder_pid") = 0
    If Len(artifact) > 0 And (Not artChanged) Then
        Dim holderBpid As Long
        holderBpid = ArtifactHolderPid(artifact)
        Dim bWriteLocked As Boolean
        bWriteLocked = ArtifactWriteLocked(artifact)
        If holderBpid > 0 Or bWriteLocked Then
            root("artifact_locked") = True
            root("artifact_locker_pid") = holderBpid
            root("artifact_holder_pid") = holderBpid
            If holderBpid > 0 Then
                root("artifact_locked_hint") = "the artifact was not rewritten and its image is running as pid " & CStr(holderBpid)
            Else
                root("artifact_locked_hint") = "the artifact was not rewritten and an exclusive open of it failed, so another process holds it (the OS names no lock owner: artifact_locker_pid = 0 means UNKNOWN, not unlocked)"
            End If
            root("artifact_locked_advice") = "do NOT just retry: rename the existing artifact first (e.g. <name>.locked-build) or close the process holding it, then rebuild"
        End If
    End If
    ' ---- TASK-P1-132 (A10): LAA on the artifact --------------------------------
    If laaUsable Then
        ApplyLaaInto root, artifact, laaWant
    Else
        LaaConfigRefusal root, laaProjFile, laaCfgErr
    End If
    root("laaSource") = laaSrc
    root("laaProjectFile") = laaProjFile
    root("laaSourceNote") = laaSrcNote
    root("laaConfigError") = laaCfgErr
    If useIde Then root("laaNote") = "the route was the in-IDE Make (see ideMake), which already ran the LAA step; this receipt re-read the artifact and reports its real state"
    ToolBuild = JsonStringify(root)
End Function

Private Function ToolEditByAnchor(ByVal name As String, ByVal anchor As String, ByVal newText As String, _
                                 ByVal expectedSha As String, ByVal occurrence As Long, ByVal matchMode As String, _
                                 ByVal scopeObj As Object, ByVal validateMode As String, ByVal snapshotOn As Boolean) As String
    Dim cur As String, nx As String, em As String
    ' TASK-P1-057: the same segment timing replace_code reports, so the ~10 s can be
    ' attributed instead of guessed.
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    Dim hitLine As Long, matched As String, rsn As String, cnt As Long
    ' TASK-P1-126 (MCP-07): how many pane lines the anchor matched (1 = single line).
    Dim spanLines As Long
    ' TASK-P1-097: one op, so no batch index is built; the candidate list is built only in
    ' the failure branch below (it costs a whole-text split + a ProcOfLine per hit).
    Dim hitsA() As Long
    Dim eRaw() As String, eNorm() As String
    Dim eFirst As Object, eCnts As Object, eRawFirst As Object, eRawCnts As Object
    If Not ResolveAnchor(name, cur, anchor, matchMode, scopeObj, occurrence, (occurrence <= 0), hitLine, matched, rsn, hitsA, cnt, spanLines, eRaw, eNorm, eFirst, eCnts, eRawFirst, eRawCnts) Then
        Dim root As Object
        Set root = CreateObject("Scripting.Dictionary")
        root("ok") = False
        root("reason") = rsn
        root("component") = name
        root("anchor") = anchor
        root("occurrence") = occurrence
        root("hitCount") = cnt
        Set root("hits") = HitList(name, cur, hitsA, cnt)
        ' ruling 4: hits[].line are line numbers, so this failure carries coords too
        Set root("coords") = CoordsDictOfName(name, 1, "ide")
        root("lineSpace") = "ide"
        root("lineSpaceHint") = "hits[].line are IDE code-pane lines; get_component_code's numbered output is DISK lines (ideLine = diskLine - coords.headerOffset)"
        root("note") = "the component was NOT modified"
        ToolEditByAnchor = JsonStringify(root)
        Exit Function
    End If
    Dim curLines() As String
    curLines = TextLines(cur)
    ' TASK-P1-126 (MCP-06/MCP-04): the same PRE-WRITE structural gate.
    Dim pairRefE As String
    pairRefE = PreWritePairingRefusal("edit_by_anchor", name, newText, "new_text")
    If Len(pairRefE) > 0 Then
        ToolEditByAnchor = pairRefE
        Exit Function
    End If
    ' ---- TASK-P1-020 beta-2 ----
    If IsContinuationTail(curLines, hitLine) Then
        ToolEditByAnchor = GuardReject(name, "edit_by_anchor", "inside_continuation", _
            "the anchor line " & CStr(hitLine) & " is the tail of a logical line: line " & CStr(hitLine - 1) & " ends with a VB continuation _. Replacing it would split that logical line; anchor the FIRST line of the logical statement instead.", _
            hitLine, hitLine, LineCount(cur))
        Exit Function
    End If
    ' TASK-P1-126 (MCP-07): a MULTI-LINE anchor replaces the WHOLE matched block, not
    ' only its first line (the A batch answered anchor_not_found for a pasted block).
    ' The extra matched lines are removed with D ops; ApplyLineOps sorts DESCENDING, so
    ' the trailing D ops are applied before the R at the block's first line.
    Dim lineOps() As String
    Dim nOps As Long
    nOps = spanLines
    If nOps < 1 Then nOps = 1
    ReDim lineOps(0 To nOps - 1)
    lineOps(0) = "R" & vbTab & CStr(hitLine) & vbTab & Replace(newText, vbCrLf, Chr$(1))
    Dim q As Long
    For q = 1 To nOps - 1
        lineOps(q) = "D" & vbTab & CStr(hitLine + q)
    Next q
    nx = EditTxn.ApplyLineOps(cur, lineOps, em)
    tPredict = GetTickCount()
    If Len(em) > 0 Then Err.Raise vbObjectError + 82, "CodeTools", em
    ' ---- TASK-P1-020 beta-3 + TASK-P1-024 gamma-4 ----
    Dim pack As Object
    ' TASK-P1-054: region target = the anchor line that was just edited.
    Set pack = LintPackForEdit(nx, KindOfName(name), LineCount(cur), hitLine)
    tLint = GetTickCount()
    Dim sref As String
    sref = StrictRefusal(name, "edit_by_anchor", pack, validateMode)
    If Len(sref) > 0 Then
        ToolEditByAnchor = sref
        Exit Function
    End If
    ' TASK-P1-087: opt-in snapshot. This tool already snapshots through the write funnel
    ' (default behaviour, unchanged); when the caller asks explicitly we take it HERE and
    ' tell the funnel to skip, so snapshot=true still writes exactly ONE file.
    If snapshotOn Then UndoSnapshot name, cur
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx, snapshotOn)
    tWrite = GetTickCount()
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    AddWritePhases extra
    Dim rq As Object
    Set rq = CreateObject("Scripting.Dictionary")
    rq("mode") = "anchor"
    rq("anchor") = anchor
    rq("occurrence") = occurrence
    rq("match") = IIf(Len(Trim$(matchMode)) = 0, "normalized", matchMode)
    If Not scopeObj Is Nothing Then Set rq("scope") = scopeObj
    Set extra("requestedRange") = rq
    Set extra("appliedRange") = AppliedRangeDict(hitLine, hitLine + nOps - 1)
    extra("clamped") = False
    extra("matchedText") = matched
    extra("matchedLine") = hitLine
    extra("matchedSpanLines") = nOps
    ' TASK-P1-126 (MCP-07): echo the TARGET PROCEDURE first/last line, before and
    ' after, computed from the pane text on both sides - so a line-number drift shows
    ' up in the receipt itself instead of silently editing another procedure.
    Set extra("procBefore") = ProcEchoDict(cur, hitLine)
    Set extra("procAfter") = ProcEchoDict(nx, hitLine)
    ToolEditByAnchor = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

' TASK-P1-020 alpha: the candidate lines of a rejected anchor, with the enclosing
' procedure and the raw line text, so the caller can pick one instead of guessing.
' TASK-P1-097: called ONLY when a failure must be reported (the success path used to build
' this for every op). procMemo caches CodeModule.ProcOfLine per line for the duration of
' one tool call, so a repeated line is never probed twice.
Private Function HitList(ByVal name As String, ByVal text As String, ByRef hits() As Long, ByVal cnt As Long, Optional ByVal procMemo As Object = Nothing) As Object
    Dim arr As Object
    Set arr = New Collection
    Dim lines() As String
    lines = TextLines(text)
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    Dim i As Long, ln As Long
    For i = 0 To cnt - 1
        MaybePump
        ln = hits(i)
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("line") = ln
        d("proc") = ProcNameOfLineCached(comp, ln, procMemo)
        If ln >= 1 And ln <= UBound(lines) + 1 Then d("text") = lines(ln - 1)
        arr.Add d
    Next
    Set HitList = arr
End Function

Private Function ProcNameOfLine(ByVal comp As VBIDE.VBComponent, ByVal line As Long) As String
    On Error Resume Next
    ProcNameOfLine = ""
    If comp Is Nothing Then Exit Function
    If comp.CodeModule Is Nothing Then Exit Function
    ProcNameOfLine = comp.CodeModule.ProcOfLine(line, vbext_pk_Proc)
End Function

' TASK-P1-097: ProcOfLine is a COM call; remember it per line for the duration of one tool
' call. Errors inside ProcNameOfLine already collapse to "" (see there), so an error entry
' is cached as "" - the same value the uncached call would return.
Private Function ProcNameOfLineCached(ByVal comp As VBIDE.VBComponent, ByVal line As Long, ByVal procMemo As Object) As String
    If procMemo Is Nothing Then
        ProcNameOfLineCached = ProcNameOfLine(comp, line)
        Exit Function
    End If
    If procMemo.Exists(line) Then
        ProcNameOfLineCached = procMemo(line)
        Exit Function
    End If
    Dim p As String
    p = ProcNameOfLine(comp, line)
    procMemo(line) = p
    ProcNameOfLineCached = p
End Function

Private Function ToolApplyEdits(ByVal name As String, ByVal edits As Object, ByVal expectedSha As String, _
                                ByVal matchMode As String, ByVal validateMode As String, ByVal snapshotOn As Boolean) As String
    Dim cur As String, nx As String, em As String, ops() As String
    Dim engTagA As String
    ' TASK-P1-057: segment timing (same shape as replace_code).
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    ' TASK-P1-094: same up-front refusal as preview_edit (before ANY resolve or write).
    Dim tooBigA As String
    tooBigA = AnchorBatchTooLarge("apply_edits", name, CollectionCount(edits), LineCount(cur))
    If Len(tooBigA) > 0 Then
        ToolApplyEdits = tooBigA
        Exit Function
    End If
    ops = BuildOps(edits)
    ' TASK-P1-126 (MCP-06/MCP-04): the same PRE-WRITE pairing gate, op by op.
    Dim pj As Long, pp() As String, ptxt As String, pr As String
    For pj = LBound(ops) To UBound(ops)
        pp = Split(ops(pj), vbTab, 3)
        ptxt = ""
        If UBound(pp) >= 2 Then ptxt = OpUnesc(pp(2))
        If Len(ptxt) > 0 Then
            pr = PreWritePairingRefusal("apply_edits", name, ptxt, "op " & CStr(pj))
            If Len(pr) > 0 Then
                ToolApplyEdits = pr
                Exit Function
            End If
        End If
    Next pj
    ' gamma-5/6/7: resolve EVERY op (occurrence + match) to a concrete line BEFORE
    ' anything is written, and keep the beta-2 continuation refusal on the resolved
    ' line. Then the whole batch is applied once through the line-addressed engine, so
    ' apply_edits stays all-or-nothing.
    Dim n As Long, i As Long, parts() As String
    n = UBound(ops) - LBound(ops) + 1
    Dim lineOps() As String
    Dim firstHit As Long
    ' TASK-P1-126 (MCP-07): a Collection, because a MULTI-LINE anchor removes a BLOCK:
    ' the extra matched lines need their own D ops, so the op count is no longer n.
    Dim colOps As Object
    Set colOps = New Collection
    Dim matches As Object
    Set matches = New Collection
    Dim curLines() As String
    curLines = TextLines(cur)
    ' TASK-P1-097: ONE index for the WHOLE batch. ResolveAnchor used to do three whole-text
    ' passes per op (VbAnchor.FindHitsEx + HitList's TextLines + matchedText's TextLines)
    ' plus one CodeModule.ProcOfLine COM call per hit, i.e. O(s x N) for s ops. The index
    ' answers a unique anchor in O(1); the HitList of a rejected op is built in the failure
    ' branch only. The exact-match index is built lazily and only if some op asks for it.
    Dim iRaw() As String, iNorm() As String
    Dim iFirst As Object, iCnts As Object, iRawFirst As Object, iRawCnts As Object
    If Len(cur) > 0 Then VbAnchor.BuildNormIndex cur, iRaw, iNorm, iFirst, iCnts
    Dim procMemo As Object
    Set procMemo = CreateObject("Scripting.Dictionary")
    Dim hitsA() As Long
    For i = 0 To n - 1
        MaybePump
        parts = Split(ops(i), vbTab)
        Dim op As String, anc As String, txt As String
        op = parts(0)
        ' TASK-P1-082: undo the TAB escape (OpEsc in BuildOps). Without this a TAB in
        ' the text shifted the fields and the occurrence conversion below raised the
        ' undiagnosable "类型不匹配".
        anc = OpUnesc(parts(1))
        txt = ""
        If UBound(parts) >= 2 Then txt = OpUnesc(parts(2))
        Dim occ As Long, mch As String
        occ = 0
        mch = matchMode
        If UBound(parts) >= 3 Then occ = CLng("0" & parts(3))
        If UBound(parts) >= 4 Then
            If Len(Trim$(parts(4))) > 0 Then mch = OpUnesc(parts(4))
        End If
        Dim hitLine As Long, matched As String, rsn As String, cnt As Long
        Dim spanLinesA As Long
        If Not ResolveAnchor(name, cur, anc, mch, Nothing, occ, (occ <= 0), hitLine, matched, rsn, hitsA, cnt, spanLinesA, iRaw, iNorm, iFirst, iCnts, iRawFirst, iRawCnts) Then
            Dim bd As Object
            Set bd = CreateObject("Scripting.Dictionary")
            bd("ok") = False
            bd("reason") = rsn
            bd("component") = name
            bd("failedOpIndex") = i
            bd("anchor") = anc
            bd("occurrence") = occ
            bd("hitCount") = cnt
            Set bd("hits") = HitList(name, cur, hitsA, cnt, procMemo)
            ' ruling 4: the failing op maps to a line, so carry coords + the space
            Set bd("coords") = CoordsDictOfName(name, 1, "ide")
            bd("lineSpace") = "ide"
            bd("lineSpaceHint") = "hits[].line are IDE code-pane lines; get_component_code's numbered output is DISK lines (ideLine = diskLine - coords.headerOffset)"
            bd("note") = "the batch was NOT applied (apply_edits is all-or-nothing)"
            ToolApplyEdits = JsonStringify(bd)
            Exit Function
        End If
        If op = "R" Or op = "D" Then
            If IsContinuationTail(curLines, hitLine) Then
                ToolApplyEdits = GuardReject(name, "apply_edits", "inside_continuation", _
                    "op " & CStr(i) & " (" & op & ") targets line " & CStr(hitLine) & ", which is the tail of a logical line: line " & CStr(hitLine - 1) & " ends with a VB continuation _. Widen the range or anchor the first line of the statement.", _
                    hitLine, hitLine, LineCount(cur))
                Exit Function
            End If
        ElseIf op = "I" Then
            If IsContinuationTail(curLines, hitLine) Then
                ToolApplyEdits = GuardReject(name, "apply_edits", "inside_continuation", _
                    "op " & CStr(i) & " (I) inserts after line " & CStr(hitLine) & ", which ends with a VB continuation _, so the insertion would split that logical line.", _
                    hitLine, hitLine, LineCount(cur))
                Exit Function
            End If
        End If
        If firstHit = 0 Then firstHit = hitLine
        colOps.Add op & vbTab & CStr(hitLine) & vbTab & txt
        ' TASK-P1-126 (MCP-07): the rest of a multi-line anchor match is REMOVED.
        If op = "R" And spanLinesA > 1 Then
            Dim dq As Long
            For dq = 1 To spanLinesA - 1
                colOps.Add "D" & vbTab & CStr(hitLine + dq)
            Next dq
        End If
        Dim md As Object
        Set md = CreateObject("Scripting.Dictionary")
        md("opIndex") = i
        md("line") = hitLine
        md("text") = matched
        matches.Add md
    Next
    ' TASK-P1-126 (MCP-07): materialise the (possibly longer) op list once.
    Dim cN As Long, cI As Long
    cN = colOps.Count
    If cN < 1 Then cN = 1
    ReDim lineOps(0 To cN - 1)
    For cI = 0 To cN - 1
        lineOps(cI) = CStr(colOps(cI + 1))
    Next cI
    nx = EditTxn.ApplyLineOps(cur, lineOps, em)
    tPredict = GetTickCount()
    ' TASK-P1-100: ApplyLineOps receives RESOLVED line ops, so its only caller-input errMsg is
    ' "unknown op X" (the anchor-shaped failures never get here: apply_edits resolves every op
    ' through ResolveAnchor first and answers those structurally already). Everything else is an
    ' engine invariant and keeps this bare raise.
    If Len(em) > 0 Then
        engTagA = EngineRefusalTag(em, ops, cur, matchMode)
        If Len(engTagA) > 0 Then Err.Raise vbObjectError + 84, "CodeTools", engTagA
        Err.Raise vbObjectError + 84, "CodeTools", em
    End If
    ' ---- TASK-P1-020 beta-3 + TASK-P1-024 gamma-4 ----
    Dim pack As Object
    ' TASK-P1-054: region target = the first edited line of the batch.
    Set pack = LintPackForEdit(nx, KindOfName(name), LineCount(cur), firstHit)
    tLint = GetTickCount()
    Dim asref As String
    asref = StrictRefusal(name, "apply_edits", pack, validateMode)
    If Len(asref) > 0 Then
        ToolApplyEdits = asref
        Exit Function
    End If
    ' TASK-P1-087: opt-in snapshot. Same shape as edit_by_anchor: explicit call here, the
    ' funnel is told to skip, so snapshot=true still writes exactly ONE file.
    If snapshotOn Then UndoSnapshot name, cur
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx, snapshotOn)
    tWrite = GetTickCount()
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    AddWritePhases extra
    ' TASK-P1-102: report every per-op key this tool cannot honour (it reads
    ' op/anchor/text plus match/occurrence) instead of ignoring it without a word.
    OpKeyAuditInto extra, edits, "|op|anchor|text|match|occurrence|", NoteApplyEditsOpKeys()
    Dim rq As Object
    Set rq = CreateObject("Scripting.Dictionary")
    rq("mode") = "anchor-batch"
    rq("ops") = CollectionCount(edits)
    rq("match") = IIf(Len(Trim$(matchMode)) = 0, "normalized", matchMode)
    Set extra("requestedRange") = rq
    Set extra("appliedRange") = ChangedRegion(cur, nx)
    extra("clamped") = False
    Set extra("matchedLines") = matches
    If matches.Count > 0 Then
        extra("matchedLine") = matches(1)("line")
        extra("matchedText") = matches(1)("text")
    End If
    ToolApplyEdits = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

' TASK-P1-020 beta-2 for the anchor batch: every op must land outside a VB
' continuation. Returns True plus the failing 0-based op index and an explanation.
Private Function ContinuationProblemOps(ByVal text As String, ByRef ops() As String, ByRef badOp As Long, ByRef detail As String) As Boolean
    Dim lines() As String
    lines = TextLines(text)
    Dim i As Long
    For i = LBound(ops) To UBound(ops)
        MaybePump
        Dim parts() As String
        parts = Split(ops(i), vbTab)
        If UBound(parts) >= 1 Then
            Dim o As String, anchor As String
            o = parts(0)
            anchor = OpUnesc(parts(1))
            Dim hits() As Long, cnt As Long
            VbAnchor.FindHits text, anchor, hits, cnt
            If cnt >= 1 Then
                Dim ln As Long
                ln = hits(0)
                If o = "R" Or o = "D" Then
                    If IsContinuationTail(lines, ln) Then
                        badOp = i
                        detail = "op " & CStr(i) & " (" & o & ") targets line " & CStr(ln) & ", the tail of a logical line (line " & CStr(ln - 1) & " ends with a VB continuation _)"
                        ContinuationProblemOps = True
                        Exit Function
                    End If
                ElseIf o = "I" Then
                    If ln >= 1 And ln <= UBound(lines) + 1 Then
                        If EndsWithCont(lines(ln - 1)) Then
                            badOp = i
                            detail = "op " & CStr(i) & " (I) inserts after line " & CStr(ln) & ", which ends with a VB continuation _, so the insertion would split that logical line"
                            ContinuationProblemOps = True
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    Next
    ContinuationProblemOps = False
End Function

Private Function EndsWithCont(ByVal line As String) As Boolean
    Dim t As String
    t = RTrim$(Vb6Lint.StripCode(line))
    If Len(t) < 2 Then Exit Function
    If Right$(t, 1) <> "_" Then Exit Function
    Dim pc As String
    pc = Mid$(t, Len(t) - 1, 1)
    If pc = " " Or pc = vbTab Then EndsWithCont = True
End Function

' ---------------------------------------------------------------------------
' TASK-P1-100: preview_edit / apply_edits answered a NON-RESOLVING ANCHOR with a bare
' Err.Raise vbObjectError + 84. CallToolInner can only wrap that as
' {ok:false, reason:"unexpected_error", vbErrorCode:-2147221420} - a value the caller cannot
' branch on, although "the anchor I passed is not in the pane" is ordinary CALLER INPUT.
'
' EditTxn.bas contains NO Err.Raise at all: ApplyEdits / ApplyLineOps report every failure as a
' plain errMsg string. This function is the ONE place that decides which errMsg values are
' caller input (-> a structured refusal through the MCPREASON channel) and which are engine
' invariants (-> the historical bare raise, byte for byte).
'
'   caller input : "op N failed: anchor not found"      -> anchor_not_found
'                  "op N failed: anchor not unique (C)" -> anchor_not_unique
'                  "unknown op X"                       -> unknown_op
'   invariants   : "bad op at N", "line N out of range (1..M)", "sha mismatch"
'                  (TASK-P1-100 section 3: none of the three can be produced by a record BuildOps
'                   emits, by a RESOLVED line set, or by preview_edit, which passes
'                   expected_sha256 = ""; they stay a bare raise on purpose and are documented)
'
' The tag carries NUMBERS ONLY (reason|opIndex|hitCount|containsHits|firstContainsLine|paneLines|
' anchorHasCrLf): an anchor can contain a "|", so it never travels through the "|"-separated
' channel - McpReasonJson reads it back out of the caller's OWN args instead.
'
' Every diagnostic runs INSIDE the failure branch only (never on the success path). Array
' subscripts are always bounds-checked on their own line first: VB6 And/Or do not short-circuit
' (TASK-P1-099), so "check the bound" and "take the element" must never share one expression.
Private Function EngineRefusalTag(ByVal em As String, ByRef ops() As String, ByVal cur As String, Optional ByVal defMatch As String = "") As String
    EngineRefusalTag = ""
    Dim reason As String
    Dim idx As Long
    idx = -1
    If EngineOpFailure(em, "anchor not found", idx) Then
        reason = "anchor_not_found"
    ElseIf EngineOpFailure(em, "anchor not unique", idx) Then
        reason = "anchor_not_unique"
    ElseIf Left$(em, 11) = "unknown op " Then
        reason = "unknown_op"
        idx = EngineOpIndex(ops, Mid$(em, 12))
    Else
        Exit Function
    End If
    ' ---- diagnostics for the ONE failing op ----
    Dim anc As String
    Dim recLine As String
    recLine = EngineOpRecord(ops, idx)
    Dim parts() As String
    ' TASK-P1-104: the failing op's OWN per-op match wins over the tool's default, so the hit
    ' count reported below is the count the resolver REALLY used. Diagnosing a per-op
    ' match="exact" op with the NORMALIZED count (and then claiming a case difference cannot
    ' be the cause) would declare an un-triggerable cause - discipline 39.
    Dim effMch As String
    effMch = defMatch
    If Len(recLine) > 0 Then
        parts = Split(recLine, vbTab)
        If LineUbound(parts) >= 1 Then anc = OpUnesc(parts(1))
        If LineUbound(parts) >= 4 Then
            If Len(Trim$(parts(4))) > 0 Then effMch = OpUnesc(parts(4))
        End If
    End If
    Dim hits() As Long, hitN As Long
    VbAnchor.FindHitsEx cur, anc, effMch, 0, 0, hits, hitN
    Dim lines() As String
    lines = TextLines(cur)
    Dim paneN As Long
    paneN = LineUbound(lines) + 1
    If paneN < 0 Then paneN = 0
    Dim containsN As Long, firstContains As Long, i As Long
    Dim na As String
    containsN = 0
    firstContains = 0
    ' TASK-P1-104: the fragment scan uses the SAME comparator as the hit scan, so under
    ' match="exact" it stays a RAW binary compare and cannot report a case-differing line as
    ' a "fragment" under a normalization that was not applied.
    If LCase$(Trim$(effMch)) = "exact" Then
        For i = 0 To LineUbound(lines)
            MaybePump
            If InStr(1, lines(i), anc, vbBinaryCompare) > 0 Then
                containsN = containsN + 1
                If firstContains = 0 Then firstContains = i + 1
            End If
        Next
    Else
        na = VbAnchor.NormLine(anc)
        If Len(na) > 0 Then
            For i = 0 To LineUbound(lines)
                MaybePump
                If InStr(1, VbAnchor.NormLine(lines(i)), na, vbBinaryCompare) > 0 Then
                    containsN = containsN + 1
                    If firstContains = 0 Then firstContains = i + 1
                End If
            Next
        End If
    End If
    Dim hasCrLf As Long
    hasCrLf = 0
    If InStr(anc, vbCr) > 0 Then hasCrLf = 1
    If InStr(anc, vbLf) > 0 Then hasCrLf = 1
    EngineRefusalTag = "MCPREASON:" & reason & "|" & CStr(idx) & "|" & CStr(hitN) & "|" _
        & CStr(containsN) & "|" & CStr(firstContains) & "|" & CStr(paneN) & "|" & CStr(hasCrLf)
End Function

' "op <digits> failed: <tail>" -> the 0-based op index. Hand-parsed (no regex object): the digits
' are validated one character at a time, and the index is only accepted when Err stayed clear.
Private Function EngineOpFailure(ByVal em As String, ByVal tail As String, ByRef idx As Long) As Boolean
    EngineOpFailure = False
    idx = -1
    If Len(em) < 8 Then Exit Function
    If Left$(em, 3) <> "op " Then Exit Function
    Dim p As Long
    p = InStr(4, em, " failed: ")
    If p = 0 Then Exit Function
    Dim s As String, k As Long, ch As String
    s = Mid$(em, 4, p - 4)
    If Len(s) = 0 Then Exit Function
    For k = 1 To Len(s)
        ch = Mid$(s, k, 1)
        If ch < "0" Or ch > "9" Then Exit Function
    Next
    Dim rest As String
    rest = Mid$(em, p + 9)
    If InStr(1, rest, tail, vbBinaryCompare) <> 1 Then Exit Function
    On Error Resume Next
    idx = CLng(s)
    If Err.Number <> 0 Then
        Err.Clear
        idx = -1
        Exit Function
    End If
    On Error GoTo 0
    EngineOpFailure = True
End Function

' the 0-based index of the first op record whose op field equals opStr, or -1.
Private Function EngineOpIndex(ByRef ops() As String, ByVal opStr As String) As Long
    EngineOpIndex = -1
    Dim i As Long, u As Long, parts() As String
    u = LineUbound(ops)
    For i = 0 To u
        If Len(Trim$(ops(i))) > 0 Then
            parts = Split(ops(i), vbTab)
            If LineUbound(parts) >= 0 Then
                If parts(0) = opStr Then
                    EngineOpIndex = i
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' ops(idx) with the bound checked FIRST (never "idx <= UBound And ops(idx)").
Private Function EngineOpRecord(ByRef ops() As String, ByVal idx As Long) As String
    EngineOpRecord = ""
    If idx < 0 Then Exit Function
    If idx > LineUbound(ops) Then Exit Function
    EngineOpRecord = ops(idx)
End Function

' ---------------------------------------------------------------------------
' TASK-P1-102: per-op keys the caller supplied but this tool does NOT honour.
'
' The DECLARED per-op shape of an anchor batch is exactly {op, anchor, text}
' (docs/tools_schema.json:1498 for apply_edits, :1540 for preview_edit; CONTRACTS.md's
' B+C table declares match/occurrence/scope on edit_by_anchor only). On top of that
' declared shape BOTH apply_edits and preview_edit HONOUR per-op match and occurrence
' (TASK-P1-104: preview_edit used to resolve through EditTxn.ApplyEdits, which read
' fields 0..2 only, so the two tools disagreed in BOTH directions), and NO tool reads a
' per-op scope (BuildOps never looks at that key at all).
'
' Silently ignoring a key the caller believes took effect changes WHERE an edit lands:
' the caller asks for "the 2nd hit" or "exact case" and quietly gets the default
' semantics. So every key outside the honoured set is REPORTED in the receipt instead.
' A call that stays inside the declared shape gets exactly the historical key set --
' nothing is added -- which is why the success path stays byte-for-byte identical.
'
' honoured is pipe-delimited with a leading and a trailing pipe ("|op|anchor|text|");
' the comparison is BINARY, matching the case sensitivity of Scripting.Dictionary and
' DictStr. Bounds are taken on their own statements first (VB6 And/Or do not
' short-circuit - TASK-P1-099 discipline 34).
Private Function OpKeyAuditRows(ByVal edits As Object, ByVal honoured As String) As Object
    Dim rows As Object, it As Variant, d As Object, k As Variant
    Dim keys As Collection, rw As Object, klist() As String
    Dim idx As Long, j As Long
    Set rows = New Collection
    idx = 0
    For Each it In edits
        If IsObject(it) Then
            Set d = it
            Set keys = New Collection
            For Each k In d.Keys
                If InStr(1, honoured, "|" & CStr(k) & "|", vbBinaryCompare) = 0 Then keys.Add CStr(k)
            Next
            If keys.Count > 0 Then
                Set rw = CreateObject("Scripting.Dictionary")
                rw("opIndex") = idx
                ReDim klist(0 To keys.Count - 1)
                For j = 1 To keys.Count
                    klist(j - 1) = keys(j)
                Next
                rw("keys") = klist
                rows.Add rw
            End If
        End If
        idx = idx + 1
    Next
    Set OpKeyAuditRows = rows
End Function

' TASK-P1-102: merge the audit into a receipt dictionary -- and ONLY when there is
' something to report. With an input inside the declared shape the receipt keeps its
' historical key set byte for byte.
Private Sub OpKeyAuditInto(ByVal root As Object, ByVal edits As Object, ByVal honoured As String, ByVal note As String)
    Dim rows As Object, it As Variant, total As Long, u As Long, l As Long
    Set rows = OpKeyAuditRows(edits, honoured)
    If rows.Count = 0 Then Exit Sub
    total = 0
    For Each it In rows
        l = LBound(it("keys"))
        u = UBound(it("keys"))
        total = total + u - l + 1
    Next
    Set root("ignoredOpKeys") = rows
    root("ignoredOpKeyCount") = total
    root("opKeysHonoured") = Split(Mid$(honoured, 2, Len(honoured) - 2), "|")
    root("ignoredOpKeysNote") = note
End Sub

Private Function NotePreviewOpKeys() As String
    NotePreviewOpKeys = "preview_edit honours the DECLARED per-op keys {op, anchor, text} plus a per-op match and occurrence - the " & _
        "SAME set apply_edits honours. Since TASK-P1-104 both tools resolve through one pipeline (one norm index, ResolveAnchor, the " & _
        "beta-2 continuation guard), so a preview taken here predicts the write those tools perform. A per-op scope is read by NO tool " & _
        "at all (BuildOps never looks at the key), so it had no effect: the anchor was resolved over the WHOLE module. For a " & _
        "range-limited edit use edit_by_anchor, which declares match/occurrence/scope."
End Function

Private Function NoteApplyEditsOpKeys() As String
    NoteApplyEditsOpKeys = "apply_edits honours the DECLARED per-op keys {op, anchor, text} plus a per-op match and occurrence; " & _
        "nothing else is read from an op. A per-op scope is read by NO tool, so it had no effect here: the anchor was resolved over " & _
        "the WHOLE module. For a range-limited edit use edit_by_anchor, which declares match/occurrence/scope."
End Function

Private Function ToolPreviewEdit(ByVal name As String, ByVal edits As Object, ByVal matchMode As String) As String
    Dim cur As String, nx As String, em As String, ops() As String, root As Object
    Dim engTag As String
    Dim ctxLines As Long
    ctxLines = 3
    cur = CompText(name)
    ' TASK-P1-094: refuse an over-large batch BEFORE any work (structured, diagnosable).
    Dim tooBig As String
    tooBig = AnchorBatchTooLarge("preview_edit", name, CollectionCount(edits), LineCount(cur))
    If Len(tooBig) > 0 Then
        ToolPreviewEdit = tooBig
        Exit Function
    End If
    ' ---- TASK-P1-104: preview_edit walks the SAME parse as apply_edits -----------------
    ' It used to call EditTxn.ApplyEdits, which reads only record fields 0..2 of every op and
    ' always requires a UNIQUE anchor. A per-op match / occurrence was therefore NOT applied,
    ' and there was no beta-2 continuation guard. Measured on the SAME input in ONE IDE session
    ' (内部证据): per-op match="exact" gave preview SUCCESS
    ' while the write answered anchor_not_found, and per-op occurrence=2 gave preview
    ' anchor_not_unique while the write SUCCEEDED - the two tools disagreed in BOTH directions,
    ' so a preview could not predict the write. This is now the same three steps as
    ' ToolApplyEdits: ONE norm index for the whole batch, ResolveAnchor with the per-op
    ' match/occurrence, the same continuation guard, then the same line-addressed engine.
    ops = BuildOps(edits)
    Dim n As Long, i As Long, parts() As String
    n = UBound(ops) - LBound(ops) + 1
    Dim lineOps() As String
    ReDim lineOps(0 To n - 1)
    Dim curLines() As String
    curLines = TextLines(cur)
    Dim iRaw() As String, iNorm() As String
    Dim iFirst As Object, iCnts As Object, iRawFirst As Object, iRawCnts As Object
    If Len(cur) > 0 Then VbAnchor.BuildNormIndex cur, iRaw, iNorm, iFirst, iCnts
    Dim hitsA() As Long
    For i = 0 To n - 1
        MaybePump
        parts = Split(ops(i), vbTab)
        Dim op As String, anc As String, txt As String
        op = parts(0)
        anc = OpUnesc(parts(1))
        txt = ""
        If UBound(parts) >= 2 Then txt = OpUnesc(parts(2))
        Dim occ As Long, mch As String
        occ = 0
        mch = ""
        If UBound(parts) >= 3 Then occ = CLng("0" & parts(3))
        If UBound(parts) >= 4 Then
            If Len(Trim$(parts(4))) > 0 Then mch = OpUnesc(parts(4))
        End If
        Dim hitLine As Long, matched As String, rsn As String, cnt As Long
        Dim spanPre As Long
        If Not ResolveAnchor(name, cur, anc, mch, Nothing, occ, (occ <= 0), hitLine, matched, rsn, hitsA, cnt, spanPre, iRaw, iNorm, iFirst, iCnts, iRawFirst, iRawCnts) Then
            ' TASK-P1-100 + TASK-P1-104: the refusal RECEIPT SHAPE must not move, so the
            ' failure travels as the very same "op N failed: anchor not found|not unique (C)"
            ' text EditTxn.ApplyEdits produced and EngineRefusalTag still owns the mapping.
            Dim etxt As String
            etxt = "op " & CStr(i) & " failed: " & rsn
            If rsn = "anchor_not_found" Then etxt = "op " & CStr(i) & " failed: anchor not found"
            If rsn = "anchor_not_unique" Then etxt = "op " & CStr(i) & " failed: anchor not unique (" & CStr(cnt) & ")"
            engTag = EngineRefusalTag(etxt, ops, cur, mch)
            If Len(engTag) > 0 Then Err.Raise vbObjectError + 84, "CodeTools", engTag
            Err.Raise vbObjectError + 84, "CodeTools", etxt
        End If
        ' beta-2 continuation refusal: byte for byte the rule ToolApplyEdits applies. The
        ' preview used to have NO guard here, so a range that splits a continued statement
        ' previewed as OK and was then REFUSED by the write.
        If op = "R" Or op = "D" Then
            If IsContinuationTail(curLines, hitLine) Then
                ToolPreviewEdit = GuardReject(name, "preview_edit", "inside_continuation", _
                    "op " & CStr(i) & " (" & op & ") targets line " & CStr(hitLine) & ", which is the tail of a logical line: line " & CStr(hitLine - 1) & " ends with a VB continuation _. Widen the range or anchor the first line of the statement.", _
                    hitLine, hitLine, LineCount(cur))
                Exit Function
            End If
        ElseIf op = "I" Then
            If IsContinuationTail(curLines, hitLine) Then
                ToolPreviewEdit = GuardReject(name, "preview_edit", "inside_continuation", _
                    "op " & CStr(i) & " (I) inserts after line " & CStr(hitLine) & ", which ends with a VB continuation _, so the insertion would split that logical line.", _
                    hitLine, hitLine, LineCount(cur))
                Exit Function
            End If
        End If
        lineOps(i) = op & vbTab & CStr(hitLine) & vbTab & txt
    Next
    nx = EditTxn.ApplyLineOps(cur, lineOps, em)
    ' TASK-P1-100: a non-resolving anchor is CALLER INPUT and is answered above (structured).
    ' ApplyLineOps receives RESOLVED line ops here, so its only caller-input errMsg is
    ' "unknown op X"; everything else is an engine invariant and keeps the bare raise.
    If Len(em) > 0 Then
        engTag = EngineRefusalTag(em, ops, cur, "")
        If Len(engTag) > 0 Then Err.Raise vbObjectError + 84, "CodeTools", engTag
        Err.Raise vbObjectError + 84, "CodeTools", em
    End If
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = name
    ' TASK-P1-024 (contract invariant 8): diff and hunks both come from the SAME
    ' line-level LCS plan, so they can never disagree.
    Dim pl As Object
    ' TASK-P1-098: cap the diff TEXT (hunks are never capped - they are the complete
    ' machine-readable description of every change). The cap is applied WHILE the diff is
    ' built, so a megabyte the caller cannot use is never materialised. With the default
    ' cap an uncapped-size diff gets the extra diffTruncated/diffLinesShown/diffLinesTotal/
    ' diffLimitLines/diffNote keys; when nothing is cut the receipt keeps its historical
    ' key set byte for byte.
    Set pl = EditTxn.DiffPlan(cur, nx, ctxLines, PreviewDiffMaxLines())
    root("diff") = pl("diff")
    Set root("hunks") = pl("hunks")
    root("contextLines") = ctxLines
    root("contextApplied") = True
    root("algorithm") = pl("algorithm")
    If pl.Exists("note") Then root("algorithmNote") = pl("note")
    If pl.Exists("diffTruncated") Then
        root("diffTruncated") = True
        root("diffLinesShown") = pl("diffLinesShown")
        root("diffLinesTotal") = pl("diffLinesTotal")
        root("diffLimitLines") = pl("diffLimitLines")
        root("diffNote") = "the diff TEXT is capped: it shows the first " & CStr(pl("diffLinesShown")) & " of " & CStr(pl("diffLinesTotal")) & " diff lines. Every change is still fully described by hunks[] (kind/start/count/oldStart/oldCount/newStart/newCount), so nothing about WHAT would change is hidden - only lines of the rendered text are omitted. Raise VB6MCP_PREVIEW_DIFF_MAX_LINES (0 = unlimited) to get the whole text."
    End If
    root("lineDelta") = LineCount(nx) - LineCount(cur)
    Set root("coords") = CoordsDictOfName(name, 1, "ide")
    ' TASK-P1-102 + TASK-P1-104: report every per-op key this tool cannot honour. Since
    ' TASK-P1-104 preview_edit resolves through the SAME pipeline as apply_edits and therefore
    ' honours op/anchor/text plus a per-op match and occurrence; only a per-op scope is left.
    OpKeyAuditInto root, edits, "|op|anchor|text|match|occurrence|", NotePreviewOpKeys()
    ' TASK-P1-102: this tool has no top-level `match` argument at all (its declared
    ' inputSchema is {name, edits}), so a top-level match used to be dropped silently
    ' while the SIBLING apply_edits honours it -- preview and write could then disagree.
    If Len(Trim$(matchMode)) > 0 Then
        root("ignoredArgs") = Array("match")
        root("ignoredArgsNote") = "preview_edit has no top-level `match` parameter; the argument you passed was IGNORED, so this " & _
            "preview describes the DEFAULT (normalized) matching for every op that does not carry its own per-op match - a per-op " & _
            "match/occurrence IS honoured (TASK-P1-104). apply_edits and edit_by_anchor DO honour the top-level match, so a preview " & _
            "taken here with match=" & Trim$(matchMode) & " can still disagree with the write those tools perform."
    End If
    ToolPreviewEdit = JsonStringify(root)
End Function

' TASK-P1-098: how many lines of the rendered diff preview_edit returns. Measured before
' this cap existed: a 49,921-line diff (27k fixture, s=500 spread) produced a 1,277,502 B
' response, and the procedure fixture produced a 1.20 MB diff even at s=50. 2,000 lines is
' the default; 0 disables the cap. Configurable via VB6MCP_PREVIEW_DIFF_MAX_LINES (same
' style as VB6MCP_LINT_MAX_DIAG / VB6MCP_ANCHOR_MAX_WORK).
Private Function PreviewDiffMaxLines() As Long
    Dim s As String, v As Long
    Const DEF As Long = 2000
    v = DEF
    s = Trim$(Environ$("VB6MCP_PREVIEW_DIFF_MAX_LINES"))
    If Len(s) > 0 Then
        On Error Resume Next
        v = CLng(s)
        If Err.Number <> 0 Then v = DEF
        On Error GoTo 0
    End If
    If v < 0 Then v = 0
    PreviewDiffMaxLines = v
End Function

Private Function ToolWriteSource(ByVal path As String, ByVal text As String, ByVal enc As String, ByVal validateMode As String) As String
    Dim full As String, b() As Byte, root As Object
    full = ResolvePath(path)
    If Len(enc) = 0 Then
        If IsVbFileExt(full) Then enc = "ansi" Else enc = "utf-8"
    End If
    ' TASK-P1-020 beta-3: lint a VB source file BEFORE it is written (warn-only).
    Dim pack As Object
    If IsVbFileExt(full) Then
        Set pack = LintPack(text, InferKind(full))
        ' TASK-P1-024 gamma-4: refuse BEFORE the file is touched
        Dim wsref As String
        wsref = StrictRefusal(full, "write_source", pack, validateMode)
        If Len(wsref) > 0 Then
            ToolWriteSource = wsref
            Exit Function
        End If
    End If
    Encoding.EncodeText text, enc, b
    Encoding.WriteAllBytesTo full, b
    Set root = CreateObject("Scripting.Dictionary")
    root("file") = full
    root("encoding") = enc
    root("bytesWritten") = Encoding.ByteCount(b)
    root("sha256") = Sha256.Sha256HexBytes(b)
    CopyLintPack root, pack
    ToolWriteSource = JsonStringify(root)
End Function

Private Function ToolVerify(ByVal name As String, ByVal diskPath As String, ByVal baselineText As String) As String
    Dim cur As String, path As String, off As Long, kind As String
    ' TASK-P1-061: segmented timing. TASK-P1-055 was rescoped to "where do the ~3.9 s
    ' of verify_scope actually go?" - the previous round attributed it by guessing.
    Dim t0 As Long, tRead As Long, tMeta As Long, tFp As Long, tHash As Long, tVc As Long
    Dim tmLint As Long, tmDiag As Long, tmDisk As Long
    t0 = GetTickCount()
    cur = CompText(name)
    tRead = GetTickCount()
    path = DiskPathOf(name, diskPath)
    off = HeaderOffsetForPath(path, LineCount(cur))
    kind = KindOfComponent(FindComponent(name))
    tMeta = GetTickCount()
    ' Card TASK-P1-011: passing buildLogPath="" left buildOk at -1 (unknown), and an
    ' unknown build still produced overall:PASS. Resolve a REAL verdict here,
    ' preferring the TASK-P1-008 build cache (key+hash must match, i.e. the recorded
    ' build is still current); otherwise report honestly that it is unknown.
    Dim bState As Long
    Dim bSource As String
    Dim bNote As String
    Dim vcLoaded As Boolean
    Dim vcAt As Double
    vcLoaded = False
    vcAt = 0
    bState = -1
    bSource = "none"
    bNote = "no build has been recorded for this project; run build for a real compile verdict"
    On Error Resume Next
    Dim prj As VBIDE.VBProject
    Set prj = GetProject()
    If Not prj Is Nothing Then
        If Len(prj.FileName) > 0 Then
            Dim vcomps As Object
            Dim vfp As Object
            Set vfp = Build.BuildFingerprint(prj.FileName, "", vcomps)
            Dim vc As Object
            Set vc = CreateObject("Scripting.Dictionary")
            vc("key") = "": vc("hash") = "": vc("ok") = False: vc("exitCode") = 0
            vc("artifact") = "": vc("size") = 0: vc("logPath") = "": vc("at") = 0
            vc("verdict") = "": vc("okByKeyword") = False
            If Build.CacheLoad(vc) Then
                vcLoaded = True
                vcAt = CDbl(vc("at"))
                If CStr(vc("key")) = CStr(vfp("key")) Then
                    If CStr(vc("hash")) = CStr(vfp("hash")) Then
                        If CBool(vc("ok")) Then bState = 1 Else bState = 0
                        bSource = "cache-fresh"
                        bNote = ""
                    Else
                        bSource = "cache-stale"
                        bNote = "the project changed after the last recorded build; run build for a current verdict"
                    End If
                Else
                    bSource = "cache-other-project"
                    bNote = "the build cache holds a different project; run build for this one"
                End If
            End If
        End If
    End If
    On Error GoTo 0
    tFp = GetTickCount()
    ' ---- TASK-P1-020 alpha: coords + the build artifact this verdict used ----
    Dim ex As Object
    Set ex = CreateObject("Scripting.Dictionary")
    Set ex("coords") = CoordsDict(1, off, "ide", (Len(path) > 0))
    ex("ideSha256") = Sha256.Sha256HexTextUtf8(cur)
    tHash = GetTickCount()
    ex("headerOffset") = off
    If vcLoaded Then
        Dim artPath As String
        artPath = CStr(vc("artifact"))
        If Len(artPath) > 0 Then
            Dim aEx As Boolean, aSz As Double, aTm As Double
            ArtifactSnap artPath, aEx, aSz, aTm
            ex("artifactPath") = artPath
            If aEx Then ex("artifactMtime") = ArtifactMtimeText(True, aTm)
            Dim artCh As Boolean
            artCh = False
            If aEx Then artCh = Build.ArtifactBuilt(artPath, vcAt, aSz)
            ex("artifactChanged") = artCh
            If bSource = "cache-fresh" And aEx And Not artCh Then
                ex("verdict") = "compiled-no-artifact-change"
                ex("artifactLockedHint") = "the recorded build says success but the artifact on disk was not refreshed (locked, or identical bytes)"
            End If
        End If
    End If
    Dim extraJson As String
    extraJson = JsonStringify(ex)
    extraJson = Mid$(extraJson, 2, Len(extraJson) - 2)
    Dim vres As String
    vres = Verify.VerifyComponent(kind, cur, path, off, baselineText, "", name, bState, bSource, bNote, extraJson, tmLint, tmDiag, tmDisk)
    tVc = GetTickCount()
    ' TASK-P1-061: splice the segment timings into the receipt. ADDITIVE only - every
    ' verdict field above is computed exactly as before.
    Dim tseg As Object
    Set tseg = CreateObject("Scripting.Dictionary")
    tseg("totalMs") = GetTickCount() - t0
    tseg("readCompTextMs") = tRead - t0
    tseg("metaMs") = tMeta - tRead
    tseg("fingerprintMs") = tFp - tMeta
    tseg("hashMs") = tHash - tFp
    tseg("verifyComponentMs") = tVc - tHash
    tseg("lintMs") = tmLint
    tseg("diagnosticsJsonMs") = tmDiag
    tseg("diskCompareMs") = tmDisk
    Dim segJson As String
    segJson = JsonStringify(tseg)
    segJson = Mid$(segJson, 2, Len(segJson) - 2)
    If Len(vres) > 1 Then
        ToolVerify = Left$(vres, Len(vres) - 1) & "," & segJson & "}"
    Else
        ToolVerify = vres
    End If
End Function

Private Function ToolInjectTrace(ByVal name As String, ByVal proc As String, ByVal exprs As String, ByVal id As String, ByVal expectedSha As String) As String
    Dim cur As String, em As String, nx As String
    ' TASK-P1-059: the same segment timing edit_by_anchor/apply_edits report (TASK-P1-057),
    ' so a slow inject is attributed instead of guessed.
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    ' TASK-P1-032: NEVER emit code that cannot compile. A modal compile-error dialog
    ' blocks the whole IDE (end/remove_trace included), so an invisible symbol is a
    ' structured refusal decided LEXICALLY - VB6 is never asked to compile.
    Dim refusal As String
    refusal = TraceExprRefusal(name, cur, proc, exprs)
    If Len(refusal) > 0 Then
        ToolInjectTrace = refusal
        Exit Function
    End If
    ' TASK-P1-129: the emitted guard calls McpTrace/McpTraceOn, so the target project
    ' must contain modMcpTrace. Attach it HERE (before the injected text is written) and
    ' PROVE it landed - the card forbids telling the caller to import a file by hand.
    Dim modInfo As Object
    Set modInfo = CreateObject("Scripting.Dictionary")
    EnsureTraceModule modInfo
    If Not TraceGen.InjectTrace(cur, proc, exprs, id, nx, em) Then Err.Raise vbObjectError + 85, "CodeTools", em
    tPredict = GetTickCount()
    Dim pack As Object
    ' TASK-P1-054: region target = the instrumented procedure's declaration line.
    Set pack = LintPackForEdit(nx, KindOfName(name), LineCount(cur), ProcHeaderLineOf(nx, proc))
    tLint = GetTickCount()
    ' TASK-P1-059: write only the changed range (SetCompText did delete-all + insert-all,
    ' i.e. TWO whole-module reformats; P1-057 measured 15.7 s vs 31 ms on 27,000 lines).
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx)
    tWrite = GetTickCount()
    ' TASK-P1-028: say what the emitted guard needs and where the sink will be, instead
    ' of "it looks successful" while nothing can possibly be written.
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("guard") = "runtime"
    extra("expressionCheck") = "ok"
    extra("checkedExpressions") = exprs
    extra("guardLine") = "If McpTraceOn() Then"
    extra("condCompRequired") = False
    extra("condCompNote") = "the emitted guard is a RUNTIME check, so it does not depend on the project's CondComp. A '#If MCP_TRACE' block would be COMPILE-TIME: the VBIDE object model does not expose CondComp and project_meta edits the .vbp text, so an already-loaded project keeps its old in-memory CondComp and the block is compiled out (measured)."
    extra("moduleRequired") = "modMcpTrace"
    extra("moduleHint") = "inject_trace ATTACHES modMcpTrace automatically when the project does not have it (TASK-P1-129): the module source is embedded in the add-in, the header comment language is chosen UI-FIRST (the IDE's own UI language when it can be read, else the running VB6.EXE version resource, else English), and the attachment is VERIFIED by re-reading the component list. Nothing has to be imported by hand. remove_trace removes the module again together with the injected blocks."
    extra("modulePresentBefore") = modInfo("modulePresentBefore")
    extra("moduleAutoAttached") = modInfo("moduleAutoAttached")
    extra("moduleAttachAction") = modInfo("moduleAttachAction")
    If modInfo.Exists("moduleAttachMs") Then extra("moduleAttachMs") = modInfo("moduleAttachMs")
    If modInfo.Exists("ideLanguage") Then extra("ideLanguage") = modInfo("ideLanguage")
    If modInfo.Exists("ideLanguageSource") Then extra("ideLanguageSource") = modInfo("ideLanguageSource")
    If modInfo.Exists("ideMenuLanguage") Then extra("ideMenuLanguage") = modInfo("ideMenuLanguage")
    If modInfo.Exists("ideMenuLanguageSource") Then extra("ideMenuLanguageSource") = modInfo("ideMenuLanguageSource")
    If modInfo.Exists("commentLanguage") Then extra("commentLanguage") = modInfo("commentLanguage")
    If modInfo.Exists("commentLanguageSource") Then extra("commentLanguageSource") = modInfo("commentLanguageSource")
    If modInfo.Exists("commentLanguageRule") Then extra("commentLanguageRule") = modInfo("commentLanguageRule")
    If modInfo.Exists("moduleCodeLines") Then extra("moduleCodeLines") = modInfo("moduleCodeLines")
    If modInfo.Exists("moduleCodeSha256") Then extra("moduleCodeSha256") = modInfo("moduleCodeSha256")
    If modInfo.Exists("moduleCodeSha256Note") Then extra("moduleCodeSha256Note") = modInfo("moduleCodeSha256Note")
    If modInfo.Exists("moduleAttachError") Then extra("moduleAttachError") = modInfo("moduleAttachError")
    If modInfo.Exists("moduleAttachNote") Then extra("moduleAttachNote") = modInfo("moduleAttachNote")
    extra("sinkPath") = SinkPathOf()
    extra("sinkNote") = "MCP_TRACE_FILE if set, else %TEMP%\mcp_trace.jsonl"
    ' TASK-P1-031: do not let the caller discover a broken sink by running the program
    ' and finding no file. Probe the sink directory here, and check that the project can
    ' actually call the guard (modMcpTrace is what defines McpTrace/McpTraceOn).
    Dim psp As Object
    Set psp = SinkProbe()
    extra("sinkDir") = psp("sinkDir")
    extra("sinkDirExists") = psp("sinkDirExists")
    extra("sinkDirWritable") = psp("sinkDirWritable")
    If psp.Exists("sinkAdvice") Then extra("sinkAdvice") = psp("sinkAdvice")
    Dim moduleNow As Boolean
    moduleNow = ModulePresent(name)
    extra("modulePresent") = moduleNow
    extra("readyToTrace") = (moduleNow And CBool(psp("sinkDirWritable")))
    If Not moduleNow Then
        extra("moduleMissingAdvice") = "modMcpTrace is STILL missing after the automatic attach attempt, so the injected guard cannot compile: VB6 will raise a MODAL compile-error dialog (measured: Compile error / Sub or Function not defined) that blocks end and remove_trace. moduleAttachAction/moduleAttachError say why the attach failed - the failure is reported, never silently downgraded."
        extra("readyToTrace") = False
    End If
    extra("enableHint") = "set MCP_TRACE_ON=1 (or MCP_TRACE_FILE=<path>) in the environment of the VB6 process BEFORE starting it; McpTraceOn() reads the environment at call time. THREE preconditions are checked in this same receipt: (1) modulePresent - inject_trace attaches modMcpTrace automatically (TASK-P1-129) and verifies it, so a false here means the attach FAILED (see moduleAttachError) and the guard cannot compile; (2) sinkDirWritable/sinkAdvice - the sink directory must be creatable (the generated module MkDirs it, else falls back to %TEMP%\mcp_trace.jsonl and logs the reason to MCP_TRACE_ERR_FILE); (3) readyToTrace - true only when both hold."
    extra("shipNote") = "the injected blocks call McpTrace/McpTraceOn, which are defined in modMcpTrace: if you keep the calls and build/ship an EXE or DLL, that module must be compiled into it. remove_trace removes the injected blocks AND the auto-attached module."
    extra("runtimeCost") = "with the runtime guard the call is always compiled: one boolean check per traced line"
    ' TASK-P1-059: the segmented timing the card requires (read/predict/lint/write + write count).
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    AddWritePhases extra
    ToolInjectTrace = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

Private Function ToolRemoveTrace(ByVal name As String, ByVal id As String, ByVal expectedSha As String) As String
    Dim cur As String, em As String, nx As String
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    If Not TraceGen.RemoveTrace(cur, id, nx, em) Then Err.Raise vbObjectError + 85, "CodeTools", em
    tPredict = GetTickCount()
    Dim pack As Object
    Set pack = LintPack(nx, KindOfName(name))
    tLint = GetTickCount()
    ' TASK-P1-059: write only the changed range (SetCompText did delete-all + insert-all,
    ' i.e. TWO whole-module reformats; P1-057 measured 15.7 s vs 31 ms on 27,000 lines).
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx)
    tWrite = GetTickCount()
    ' TASK-P1-149: only THIS id's injected block is removed. modMcpTrace is KEPT BY DESIGN -
    ' the user may have added code to it, and other @mcp-trace blocks elsewhere may still
    ' call it, so the project would stop compiling without it. The receipt reports the state
    ' and the manual removal steps. (The TASK-P1-129 automatic removal was withdrawn.)
    Dim modInfo As Object
    Set modInfo = CreateObject("Scripting.Dictionary")
    ReportTraceModuleKept modInfo
    ' TASK-P1-031 (P4-028 section 15.2): cleaning the PANE is not enough. After a run,
    ' VB6's pre-run save has already written the injected block to the FILE, so the disk
    ' keeps a copy, the pane and the disk silently diverge and verify() reports FAIL.
    ' Flush the cleaned pane the same way set_breakpoint/clear_breakpoints do, then say
    ' honestly whether the disk now matches (ok:false + reason when it does not).
    Dim paneNow As String
    paneNow = CompText(name)
    Dim flushErr As String
    flushErr = ""
    On Error Resume Next
    ToolSaveProject ""
    If Err.Number <> 0 Then flushErr = Err.Description
    Err.Clear
    On Error GoTo 0
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("flushed") = True
    extra("saved") = True
    extra("diskAfterRemoveSha256") = ComponentFileSha(name)
    extra("diskStillHasMarker") = (InStr(1, DiskTextOfName(name), "@mcp-trace " & id) > 0)
    If Not DiskMatchesPane(name, paneNow) Then
        extra("ok") = False
        extra("saved") = False
        If Len(flushErr) > 0 Then
            extra("reason") = "disk_not_flushed"
            extra("note") = "the block is gone from the code pane but the FILE still holds it (flush failed: " & flushErr & "); call save_project, or expect verify() to report FAIL"
        Else
            extra("reason") = "disk_still_differs"
            extra("note") = "the block is gone from the code pane but the FILE still differs; call save_project, or expect verify() to report FAIL"
        End If
    End If
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    extra("modulePresentBefore") = modInfo("modulePresentBefore")
    extra("moduleRemoved") = modInfo("moduleRemoved")
    extra("moduleLeftInPlace") = modInfo("moduleLeftInPlace")
    extra("moduleAutogenMarker") = modInfo("moduleAutogenMarker")
    extra("moduleRemoveReason") = modInfo("moduleRemoveReason")
    If modInfo.Exists("moduleRemoveNote") Then extra("moduleRemoveNote") = modInfo("moduleRemoveNote")
    If modInfo.Exists("moduleKeptNote") Then extra("moduleKeptNote") = modInfo("moduleKeptNote")
    If modInfo.Exists("moduleFileOnDisk") Then extra("moduleFileOnDisk") = modInfo("moduleFileOnDisk")
    AddWritePhases extra
    ToolRemoveTrace = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

Private Function ToolFaultInject(ByVal name As String, ByVal proc As String, ByVal mode As String, ByVal value As String, ByVal id As String, ByVal expectedSha As String) As String
    Dim cur As String, em As String, nx As String
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    If Not TraceGen.InjectFault(cur, proc, mode, value, id, nx, em) Then Err.Raise vbObjectError + 85, "CodeTools", em
    tPredict = GetTickCount()
    Dim pack As Object
    ' TASK-P1-054: region target = the faulted procedure's declaration line.
    Set pack = LintPackForEdit(nx, KindOfName(name), LineCount(cur), ProcHeaderLineOf(nx, proc))
    tLint = GetTickCount()
    ' TASK-P1-059: write only the changed range (SetCompText did delete-all + insert-all,
    ' i.e. TWO whole-module reformats; P1-057 measured 15.7 s vs 31 ms on 27,000 lines).
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx)
    tWrite = GetTickCount()
    ' TASK-P1-059: segmented timing (was ok/ideTotalLines only).
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    AddWritePhases extra
    ToolFaultInject = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

Private Function ToolClearFaults(ByVal name As String, ByVal id As String, ByVal expectedSha As String) As String
    Dim cur As String, em As String, nx As String
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    If Not TraceGen.RemoveFault(cur, id, nx, em) Then Err.Raise vbObjectError + 85, "CodeTools", em
    tPredict = GetTickCount()
    Dim pack As Object
    Set pack = LintPack(nx, KindOfName(name))
    tLint = GetTickCount()
    ' TASK-P1-059: write only the changed range (SetCompText did delete-all + insert-all,
    ' i.e. TWO whole-module reformats; P1-057 measured 15.7 s vs 31 ms on 27,000 lines).
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx)
    tWrite = GetTickCount()
    ' TASK-P1-059: segmented timing (was ok/ideTotalLines only).
    Dim extra As Object
    Set extra = CreateObject("Scripting.Dictionary")
    extra("writeCalls") = wCalls
        extra("hunks") = m_lastWriteHunks
    extra("readMs") = tRead - tEntry
    extra("predictMs") = tPredict - tRead
    extra("lintMs") = tLint - tPredict
    extra("writeMs") = tWrite - tLint
    AddWritePhases extra
    ToolClearFaults = ChangeResult2(name, cur, nx, pack, extra, tEntry)
End Function

Private Function ToolAuditTail(ByVal n As Long) As String
    If n <= 0 Then n = 50
    ToolAuditTail = Audit.AuditTailJson(n)
End Function

' =====================================================================
' Contract v1.1.0 batch A (TASK-P1-017): ide_state / list_breakpoints / find_gui
' All three are READ-ONLY observers: none of them may change state.
' =====================================================================

' Contract section 1: the IDE mode, read live. Deliberately NOT cached - the whole
' point is that a design/run/break switch is visible on the very next call.
Private Function ToolIdeState(Optional ByVal cancelInflight As Boolean = False, Optional ByVal cancelTarget As String = "", Optional ByVal releaseInflight As Boolean = False, Optional ByVal releaseMinHeldMs As Long = 0) As String
    On Error Resume Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim st As String
    st = GetIdeState()
    root("state") = st
    root("ebMode") = IdeEbMode()
    root("title") = IdeMainTitle()
    root("project") = ProjectNameSafe()
    root("running") = (st = "run")
    ' TASK-P1-088: read-only wheel state next to the IDE state (pure addition; no new tool
    ' and no description change, so no re-freeze). See WinOps.WheelStateObj for the
    ' three-reason decision table.
    StampWheelNow root
    ' TASK-P1-126 (MCP-08): the cancel{} block, with a FIXED key set on every path.
    Dim cb As Object
    Set cb = CancelRequestBody("ide_state", cancelTarget, cancelInflight)
    If cb Is Nothing Then
        Set cb = CreateObject("Scripting.Dictionary")
        cb("requested") = cancelInflight
        cb("applied") = False
        cb("targetTool") = ""
        cb("targetKind") = ""
        cb("sinceMs") = 0
        cb("depth") = m_callDepth
        cb("reason") = "internal_error"
        cb("honorMinMs") = CANCEL_HONOR_MIN_MS
        cb("note") = "the cancel request could not be evaluated; nothing was changed"
    End If
    Set root("cancel") = cb
    ' TASK-P1-168 (C / defect #10): the release{} block, with a FIXED key set on every path
    ' (same discipline as cancel{} above). This read-only tool is the only channel that is
    ' still served while the gate is stuck, so it is where the forced release lives.
    Dim rb As Object
    Set rb = ReleaseRequestBody(releaseInflight, releaseMinHeldMs)
    If rb Is Nothing Then
        Set rb = CreateObject("Scripting.Dictionary")
        rb("attempted") = True
        rb("released") = False
        rb("reason") = "internal_error"
        rb("note") = "the release request could not be evaluated; nothing was changed"
    End If
    Set root("release") = rb
    ToolIdeState = JsonStringify(root)
End Function

' TASK-P1-168 (C): build the release{} block of an ide_state receipt. The KEY SET is fixed
' and identical on every path, exactly like cancel{} (discipline 98).
Private Function ReleaseRequestBody(ByVal requested As Boolean, ByVal minHeldMs As Long) As Object
    If requested Then
        If minHeldMs <= 0 Then minHeldMs = GATE_RELEASE_MIN_MS
        Set ReleaseRequestBody = ForceReleaseGate("ide_state", "forced_release", minHeldMs, 1)
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("attempted") = False
    d("released") = False
    d("releasedTool") = ""
    d("releasedKind") = ""
    d("heldMs") = 0
    d("minHeldMs") = GATE_RELEASE_MIN_MS
    d("retryAfterMs") = 0
    d("depthBefore") = m_callDepth
    d("depthAfter") = m_callDepth
    d("targetDepth") = 1
    d("mutBusyBefore") = m_mutBusy
    d("abandonedSeq") = 0
    d("reason") = "not_requested"
    d("gateReleaseCount") = m_gateReleaseCount
    d("statementInterrupted") = False
    d("timeoutMsEnforcedBy") = "gate_age_threshold"
    d("note") = "pass release_inflight=true (optionally release_min_held_ms) to FORCE-RELEASE the single-flight gate of a call that never came back. Read-only, so it is served while the gate is held."
    d("recoveryAction") = ""
    d("recoveryHint") = ""
    Set ReleaseRequestBody = d
End Function

' The IDE main window caption. RUN mode suspends the VBIDE object model, so fall
' back to the Win32 caption of this process's own main window.
Private Function IdeMainTitle() As String
    On Error Resume Next
    IdeMainTitle = ""
    Dim t As String
    t = ""
    t = m_vbe.MainWindow.Caption
    If Err.Number <> 0 Then Err.Clear
    If Len(t) = 0 Then t = WinProbe.MainWindowTitleOf(WinProbe.CurrentProcessId())
    IdeMainTitle = t
End Function

' Contract section 3: enumerate the breakpoints THIS tool writes into source. The
' VBIDE object model exposes no breakpoint collection at all, so IDE-native
' breakpoints genuinely cannot be listed - that is reported, never faked.
' Contract section 2 (TASK-P1-018): where the engine is stopped. VBIDE exposes no
' "current statement" API, so the channel is the ACTIVE CODE PANE's selection.
' Whether that selection follows the engine in break mode is a claim, so the tool
' also echoes the raw selection it read (selectedLine / selectedLineEnd): a caller
' can see the reported line IS the selection and cross-check the line text, instead
' of trusting a number. Outside break mode it answers structurally and never raises.
Private Function ToolCurrentPosition() As String
    On Error Resume Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim st As String
    st = GetIdeState()
    root("state") = st
    root("ebMode") = IdeEbMode()
    If st <> "break" Then
        root("available") = False
        root("reason") = "not_in_break"
        ToolCurrentPosition = JsonStringify(root)
        Exit Function
    End If
    Dim cp As Object
    Set cp = m_vbe.ActiveCodePane
    If cp Is Nothing Then
        root("available") = False
        root("reason") = "no_active_code_pane"
        ToolCurrentPosition = JsonStringify(root)
        Exit Function
    End If
    Dim cm As Object
    Set cm = cp.CodeModule
    Dim sl As Long, sc As Long, el As Long, ec As Long
    sl = 0: sc = 0: el = 0: ec = 0
    cp.GetSelection sl, sc, el, ec
    If Err.Number <> 0 Then
        root("available") = False
        root("reason") = "selection_unavailable"
        root("detail") = Err.Description
        Err.Clear
        ToolCurrentPosition = JsonStringify(root)
        Exit Function
    End If
    Dim n As Long
    n = cm.CountOfLines
    Dim line As Long
    line = sl
    If line < 1 Then line = 1
    If n > 0 And line > n Then line = n
    Dim proc As String
    proc = ""
    proc = cm.ProcOfLine(line, 0)
    If Err.Number <> 0 Then Err.Clear
    root("available") = True
    root("project") = ProjectNameSafe()
    root("module") = CodeModuleOwnerName(cm)
    root("proc") = proc
    root("line") = line
    root("text") = cm.Lines(line, 1)
    root("selectedLine") = sl
    root("selectedLineEnd") = el
    ToolCurrentPosition = JsonStringify(root)
End Function

' The VBComponent that owns a code module. CodeModule.Parent is the documented
' route, but it can fail while the program is stopped, so fall back to an identity
' scan of the project's components rather than guessing from a name.
Private Function CodeModuleOwnerName(ByVal cm As Object) As String
    On Error Resume Next
    CodeModuleOwnerName = ""
    Dim nm As String
    nm = ""
    ' 1) the cheapest and most reliable route: CodeModule.Name IS the module name, and
    ' it keeps working while the program is stopped (the same property the existing
    ' get_selection tool already uses for activeCodeModule). This is what the first
    ' attempt missed - cm.Parent and proxy identity both fail in break mode.
    nm = cm.Name
    If Err.Number <> 0 Then Err.Clear
    If Len(nm) > 0 Then
        CodeModuleOwnerName = nm
        Exit Function
    End If
    ' 2) the documented whole-object route
    nm = cm.Parent.Name
    If Err.Number <> 0 Then Err.Clear
    If Len(nm) > 0 Then
        CodeModuleOwnerName = nm
        Exit Function
    End If
    ' Identity comparison is unreliable while the program is STOPPED: each access to
    ' CodeModule can hand back a fresh proxy, so "Is" fails and the module came back
    ' empty. Fall back to a CONTENT fingerprint instead: only a module whose text
    ' matches exactly can be the owner, and an ambiguous match reports "" rather than
    ' guessing a name.
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim n As Long, sig As String
    n = cm.CountOfLines
    sig = ModuleSignature(cm)
    Dim hit As String, hits As Long
    hit = ""
    hits = 0
    Dim c As VBIDE.VBComponent
    For Each c In p.VBComponents
        Dim m As Object
        Set m = Nothing
        Set m = c.CodeModule
        If Not m Is Nothing Then
            If m.CountOfLines = n Then
                If ModuleSignature(m) = sig Then
                    hits = hits + 1
                    hit = c.Name
                End If
            End If
        End If
    Next
    If hits = 1 Then CodeModuleOwnerName = hit
End Function

' A cheap but discriminating fingerprint of a code module: the line count plus its
' first, middle and last lines. A false match would need a different module with the
' same length AND the same three sentinel lines, which is not a realistic ambiguity;
' anything that still looks ambiguous is reported as "" by the caller.
Private Function ModuleSignature(ByVal cm As Object) As String
    Dim n As Long
    n = cm.CountOfLines
    If n <= 0 Then
        ModuleSignature = "0"
        Exit Function
    End If
    Dim mid As Long
    mid = n \ 2
    If mid < 1 Then mid = 1
    ModuleSignature = CStr(n) & Chr$(1) & cm.Lines(1, 1) & Chr$(1) & cm.Lines(mid, 1) & Chr$(1) & cm.Lines(n, 1)
End Function

Private Function ToolListBreakpoints() As String
    On Error Resume Next
    Dim inj As Object
    Set inj = New Collection
    Dim mark As String
    mark = "' @mcp-break "
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim c As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim n As Long, i As Long, ln As String, id As String, stmt As String
    For Each c In p.VBComponents
        Set cm = c.CodeModule
        n = cm.CountOfLines
        For i = 1 To n
            ln = Trim$(cm.Lines(i, 1))
            If Len(ln) > Len(mark) + 6 Then
                If Left$(ln, Len(mark)) = mark And Right$(ln, 6) = " begin" Then
                    id = Trim$(Mid$(ln, Len(mark) + 1, Len(ln) - Len(mark) - 6))
                    stmt = ""
                    If i < n Then stmt = Trim$(cm.Lines(i + 1, 1))
                    Dim d As Object
                    Set d = CreateObject("Scripting.Dictionary")
                    d("component") = c.Name
                    d("line") = i + 1
                    d("markerLine") = i
                    d("id") = id
                    d("kind") = "injected"
                    d("enabled") = True
                    d("statement") = stmt
                    inj.Add d
                End If
            End If
        Next
    Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Set root("injected") = inj
    root("injectedCount") = inj.Count
    root("note") = "only breakpoints injected by set_breakpoint are listable: the VBIDE object model has no breakpoint API, so IDE-native breakpoints cannot be enumerated"
    Dim js As String
    js = JsonStringify(root)
    ' Make the contract shape explicit instead of depending on how this build of the
    ' JSON writer renders a Null (null / empty string / key omitted).
    If Len(js) > 0 And Right$(js, 1) = "}" Then
        js = Left$(js, Len(js) - 1) & "," & Chr$(34) & "native" & Chr$(34) & ":null}"
    End If
    ToolListBreakpoints = js
End Function

' Contract section 4: find_gui. The enumeration itself lives in the standard module
' WinProbe.bas, because an AddressOf callback is only legal in a standard module.
Private Function ToolFindGui(ByVal args As Object) As String
    On Error Resume Next
    Dim sh As Long
    sh = -1
    If Not args Is Nothing Then
        If args.Exists("showing") Then
            Dim bv As Variant
            bv = args("showing")
            If Not IsNull(bv) Then
                If Not IsObject(bv) Then
                    If CBool(bv) Then sh = 1 Else sh = 0
                End If
            End If
        End If
    End If
    ToolFindGui = WinProbe.FindGuiJson(JsonGetStr(args, "keyword"), JsonGetStr(args, "cls"), sh)
End Function

Private Function ToolSetBreakpoint(ByVal name As String, ByVal anchor As String, ByVal id As String, ByVal expectedSha As String) As String
    ToolSetBreakpoint = BreakpointEdit(name, "set", anchor, id, expectedSha)
End Function

Private Function ToolClearBreakpoints(ByVal name As String, ByVal id As String, ByVal expectedSha As String) As String
    ToolClearBreakpoints = BreakpointEdit(name, "clear", "", id, expectedSha)
End Function

' The component file's raw byte hash on disk ("" when it cannot be resolved).
Private Function ComponentFileSha(ByVal name As String) As String
    On Error Resume Next
    ComponentFileSha = ""
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    Dim f As String
    f = ""
    f = comp.FileNames(1)
    If Len(f) = 0 Then Exit Function
    ComponentFileSha = Build.FileSha256(f)
End Function

' True when the file on disk holds exactly the code pane's text, ignoring the
' trailing newline VB6 writes and the file's ANSI encoding. Comparing the two
' raw hashes would always differ (bytes vs text), so consistency is decided on
' normalised text; the raw hashes are reported separately.
' Normalised code text: line feeds unified, trailing newlines dropped, and the
' Attribute VB_* metadata lines removed. VB6 keeps e.g. "Attribute VB_Name" in the
' .bas FILE but the code module does not expose it, so modLoop.bas has 50 file
' lines while the pane reports 49 - comparing the raw texts would never match.
' TASK-P1-121 item 3 (spec: 内部证据).
' The STRUCTURAL head/body splitter. It identifies the head by TYPE/RULE, never by
' searching the disk for the pane's first line (that naive route mis-splits whenever
' the pane is empty or the first line also occurs inside the head - P4-106 measured
' it eating the body on the two empty-pane corpora). Rules:
'   .cls            VERSION <x> -> BEGIN..END (nested) -> leading contiguous Attribute lines
'   .frm/.ctl/.dsr  VERSION <x.y> -> Begin VB.<Kind>..End (nested Begin/End,
'                   BeginProperty/EndProperty) -> leading Attribute lines
'   .bas            leading contiguous Attribute lines only
' The body is the DISK SLICE from the first line after the head, byte-for-byte: an
' Attribute line or an End line INSIDE the body is never touched.
' FAILURE SAFETY (the whole point): if the pane is empty, or the derived body cannot be
' reconciled with the pane (equal / prefix either way / same first line), the caliber is
' 'unknown' and the DISK TEXT IS RETURNED UNCHANGED - 'rather do not split than eat the
' body'. verifyPane=False asks for the structural strip alone (add_code appends a FILE,
' so there is no pane of that file to validate - D-178).
Private Function StrippedDiskText(ByVal disk As String, ByVal paneText As String, _
                                  Optional ByVal verifyPane As Boolean = True, _
                                  Optional ByRef outCaliber As String = "") As String
    On Error Resume Next
    outCaliber = "unknown"
    StrippedDiskText = disk
    If Len(disk) = 0 Then Exit Function
    Dim lines() As String
    lines = Split(Replace(Replace(disk, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim n As Long
    n = UBound(lines)
    Dim i As Long, u As String
    i = 0
    ' ---- 1. the designer head -----------------------------------------
    u = UCase$(Trim$(lines(i)))
    If Left$(u, 8) = "VERSION " Then
        i = i + 1
        If i <= n Then
            u = UCase$(Trim$(lines(i)))
            If u = "BEGIN" Then
                i = ClsHeadBlockEnd(lines, i, n)
            ElseIf Left$(u, 6) = "BEGIN " Then
                i = DesignerHeadBlockEnd(lines, i, n)
            End If
        End If
    End If
    ' ---- 2. the LEADING contiguous Attribute block ---------------------
    Do While i <= n
        MaybePump
        If LCase$(Left$(Trim$(lines(i)), 10)) = "attribute " Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    ' ---- 3. body = the disk slice, byte-for-byte ----------------------
    Dim body As String
    body = ""
    If i <= n Then body = JoinRange(lines, i, n)
    ' ---- 4. failure safety -------------------------------------------
    If verifyPane Then
        Dim ok As Boolean
        ok = False
        If Len(paneText) = 0 Then
            ok = False
        ElseIf body = paneText Then
            ok = True
        ElseIf Len(body) > Len(paneText) Then
            If Left$(body, Len(paneText)) = paneText Then ok = True
        ElseIf Len(paneText) > Len(body) Then
            If Left$(paneText, Len(body)) = body Then ok = True
        End If
        If Not ok Then
            Dim bl() As String, pl() As String
            bl = Split(Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf), vbLf)
            pl = Split(Replace(Replace(paneText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
            If UBound(bl) >= 0 Then
                If UBound(pl) >= 0 Then
                    If bl(0) = pl(0) Then ok = True
                End If
            End If
        End If
        If Not ok Then
            outCaliber = "unknown"
            StrippedDiskText = disk
            Err.Clear
            Exit Function
        End If
    End If
    outCaliber = "ok"
    StrippedDiskText = body
    Err.Clear
End Function

' .cls designer block: BEGIN .. END, nesting counted.
Private Function ClsHeadBlockEnd(ByRef lines() As String, ByVal startIdx As Long, ByVal lastIdx As Long) As Long
    Dim depth As Long, i As Long, u As String
    depth = 0
    i = startIdx
    Do While i <= lastIdx
        MaybePump
        u = UCase$(Trim$(lines(i)))
        If u = "BEGIN" Then
            depth = depth + 1
        ElseIf u = "END" Then
            depth = depth - 1
            If depth <= 0 Then
                ClsHeadBlockEnd = i + 1
                Exit Function
            End If
        End If
        i = i + 1
    Loop
    ClsHeadBlockEnd = i
End Function

' .frm/.ctl/.dsr designer block: Begin .. End with nested Begin / BeginProperty.
Private Function DesignerHeadBlockEnd(ByRef lines() As String, ByVal startIdx As Long, ByVal lastIdx As Long) As Long
    Dim depth As Long, i As Long, u As String
    depth = 0
    i = startIdx
    Do While i <= lastIdx
        MaybePump
        u = UCase$(Trim$(lines(i)))
        If u = "BEGIN" Or Left$(u, 6) = "BEGIN " Or u = "BEGINPROPERTY" Or Left$(u, 14) = "BEGINPROPERTY " Then
            depth = depth + 1
        ElseIf u = "END" Or u = "ENDPROPERTY" Then
            depth = depth - 1
            If depth <= 0 Then
                DesignerHeadBlockEnd = i + 1
                Exit Function
            End If
        End If
        i = i + 1
    Loop
    DesignerHeadBlockEnd = i
End Function

Private Function NormalizeCode(ByVal text As String) As String
    ' TASK-P1-053: linear. This used to accumulate with "out = out & lines(i) & vbLf"
    ' once per line, which is O(n^2) on a BSTR: P4-040 measured 31-50 s on a
    ' 27,003-line / 1.5 MB pane, reproduced as 52,106 ms by verify_scope (it reaches
    ' here through DiskMatchesPane). The kept lines are now collected into a String()
    ' and joined ONCE, and the trailing-LF trim is done with a single Left$ instead
    ' of a per-character loop; measured 3,203 ms -> 15 ms at 27,003 lines, and the
    ' output is bit-identical.
    Dim lines() As String, i As Long, n As Long, k As Long, t As String
    Dim keep() As String, out As String, ln As Long
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    n = UBound(lines)
    ReDim keep(0 To n)
    k = 0
    For i = 0 To n
        MaybePump
        t = Trim$(lines(i))
        If LCase$(Left$(t, 10)) <> "attribute " Then
            keep(k) = lines(i)
            k = k + 1
        End If
    Next
    If k = 0 Then Exit Function
    If k <= n Then ReDim Preserve keep(0 To k - 1)
    out = Join(keep, vbLf) & vbLf
    ln = Len(out)
    Do While ln > 0 And Mid$(out, ln, 1) = vbLf
        ln = ln - 1
    Loop
    NormalizeCode = Left$(out, ln)
End Function

' TASK-P1-022: the disk text from the pane's FIRST line onwards, i.e. with the
' file's own header (designer block + Attribute lines) removed. Without this the
' comparison was structurally false for EVERY form: a .frm file starts with the
' designer block, NormalizeCode only strips Attribute lines, so the two normalised
' texts could never be equal (measured: onDisk=false on every set_breakpoint
' against Form1 while the file had in fact just been rewritten).
Private Function PaneTailOfDisk(ByVal disk As String, ByVal paneText As String) As String
    On Error Resume Next
    PaneTailOfDisk = disk
    Dim plines() As String, dlines() As String
    plines = Split(Replace(Replace(paneText, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    dlines = Split(Replace(Replace(disk, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    If UBound(plines) < 0 Then Exit Function
    Dim first As String, i As Long, j As Long, out As String
    first = plines(0)
    For i = 0 To UBound(dlines)
        MaybePump
        If dlines(i) = first Then
            ' TASK-P1-049: ONE Join of the slice instead of a per-line accumulation
            ' (quadratic on a whole module).
            PaneTailOfDisk = JoinRange(dlines, i, UBound(dlines))
            Exit Function
        End If
    Next
End Function

Private Function DiskMatchesPane(ByVal name As String, ByVal paneText As String) As Boolean
    On Error Resume Next
    DiskMatchesPane = False
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(name)
    If comp Is Nothing Then Exit Function
    Dim f As String
    f = ""
    f = comp.FileNames(1)
    If Len(f) = 0 Then Exit Function
    If Dir$(f) = "" Then Exit Function
    DiskMatchesPane = (NormalizeCode(PaneTailOfDisk(Build.ReadLogAnsi(f), paneText)) = NormalizeCode(paneText))
End Function

' Card TASK-P1-009: apply a trace-breakpoint edit and then reconcile the DISK
' with the code pane HONESTLY.
' The edit used to live only in the IDE pane. The injected block reached the file
' only as a side effect of run's pre-run flush (TASK-P1-005), so clear_breakpoints
' answered with fresh hashes while the disk still held the injection - a fake
' "cleared" state (P4: now=9ECDFD47 != backup=9FCE7D14). We now flush and report
' diskSha256 / paneSha256 / onDisk, and when the disk cannot be made to match we
' answer ok:false with a reason instead of claiming success.
' TASK-P1-022: is 'line' a statement INSIDE a procedure body? See the body of this
' function for why the VBIDE proc queries are not usable for it.
Private Function InsideProcedureBody(ByVal name As String, ByVal line As Long, ByVal paneText As String) As Boolean
    On Error Resume Next
    InsideProcedureBody = False
    If line < 1 Then Exit Function
    ' Use the lint engine's own procedure map: it is a real VB parser, whereas
    ' VBIDE's ProcOfLine attributes a blank line before a procedure to that procedure
    ' and ProcBodyLine returns the declaration line itself on this build (both
    ' measured - two earlier versions of this guard let the declaration through and
    ' the project then stopped compiling).
    Vb6Lint.Lint paneText, KindOfName(name)
    If Len(Vb6Lint.ProcBodyOwner(line)) > 0 Then InsideProcedureBody = True
End Function

Private Function BreakpointEdit(ByVal name As String, ByVal op As String, ByVal anchor As String, _
                                ByVal id As String, ByVal expectedSha As String) As String
    Dim cur As String, em As String, nx As String
    ' TASK-P1-059: segmented timing for set_breakpoint/clear_breakpoints (this receipt used to
    ' carry ok/ideTotalLines only).
    Dim tEntry As Long, tRead As Long, tPredict As Long, tLint As Long, tWrite As Long
    tEntry = GetTickCount()
    cur = CompText(name)
    GuardSha cur, expectedSha
    tRead = GetTickCount()
    If op = "clear" Then
        If Not TraceGen.ClearBreakpoint(cur, id, nx, em) Then Err.Raise vbObjectError + 88, "CodeTools", em
    Else
        ' TASK-P1-020 beta-2: never inject inside a VB line continuation
        Dim curLines() As String
        curLines = TextLines(cur)
        Dim hits() As Long, cnt As Long
        VbAnchor.FindHits cur, anchor, hits, cnt
        If cnt >= 1 Then
            If IsContinuationTail(curLines, hits(0)) Then
                BreakpointEdit = GuardReject(name, "set_breakpoint", "inside_continuation", _
                    "the anchor line " & CStr(hits(0)) & " is the tail of a logical line (line " & CStr(hits(0) - 1) & " ends with a VB continuation _); injecting there would split it. Anchor the first line of the logical statement.", _
                    hits(0), hits(0), LineCount(cur))
                Exit Function
            End If
            ' TASK-P1-022: the injected Debug.Assert False must land INSIDE a procedure.
            If Not InsideProcedureBody(name, hits(0), cur) Then
                BreakpointEdit = GuardReject(name, "set_breakpoint", "anchor_outside_procedure", _
                    "the anchor line " & CStr(hits(0)) & " is not inside a procedure body (it is a module-level line or a procedure declaration). Injecting Debug.Assert False there puts executable code OUTSIDE any procedure, and the project then stops compiling with 'invalid outside procedure'. Anchor a statement inside the procedure body instead.", _
                    hits(0), hits(0), LineCount(cur))
                Exit Function
            End If
        End If
        If Not TraceGen.SetBreakpoint(cur, anchor, id, nx, em) Then Err.Raise vbObjectError + 88, "CodeTools", em
    End If
    tPredict = GetTickCount()
    ' TASK-P1-020 beta-3: lint the RESULT before it is written (warn-only)
    Dim pack As Object
    ' TASK-P1-054: region target = the anchored line (set); clear has no line -> whole.
    Dim bpTarget As Long
    If op = "set" Then
        If cnt >= 1 Then bpTarget = hits(0)
    End If
    Set pack = LintPackForEdit(nx, KindOfName(name), LineCount(cur), bpTarget)
    tLint = GetTickCount()
    ' TASK-P1-059: write only the changed range (SetCompText did delete-all + insert-all,
    ' i.e. TWO whole-module reformats; P1-057 measured 15.7 s vs 31 ms on 27,000 lines).
    Dim wCalls As Long
    wCalls = SetCompTextMinimal(name, cur, nx)
    tWrite = GetTickCount()
    Dim paneNow As String
    paneNow = CompText(name)
    Dim paneSha As String
    paneSha = Sha256.Sha256HexTextUtf8(paneNow)
    ' flush the pane to its file; a failure is reported, never hidden
    Dim flushErr As String
    flushErr = ""
    On Error Resume Next
    ToolSaveProject ""
    If Err.Number <> 0 Then flushErr = Err.Description
    Err.Clear
    On Error GoTo 0
    Dim diskSha As String
    diskSha = ComponentFileSha(name)
    Dim matches As Boolean
    matches = DiskMatchesPane(name, paneNow)
    Dim why As String
    why = ""
    If Not matches Then
        If Len(flushErr) > 0 Then
            why = "not persisted to disk: " & flushErr
        Else
            why = "disk still differs from the code pane after the flush attempt"
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("component") = name
    root("operation") = op
    ' TASK-P1-085: the same project-directory write manifest as every other write receipt.
    Dim pdwB As Object
    Set pdwB = New Collection
    ManifestMergeUndo pdwB
    Set root("projectDirWrites") = pdwB
    root("sha256Before") = Sha256.Sha256HexTextUtf8(cur)
    root("sha256After") = paneSha
    root("lineDelta") = LineCount(paneNow) - LineCount(cur)
    root("newTotalLines") = LineCount(paneNow)
    ' Card TASK-P1-009 criterion 3: the disk/pane reconciliation fields
    root("paneSha256") = paneSha
    root("diskSha256") = diskSha
    root("onDisk") = matches
    root("ok") = (Len(why) = 0)
    If Len(why) > 0 Then root("reason") = why
    ' ---- TASK-P1-020 alpha: coords + the IDE-read line count ----
    Dim ideTotal As Long
    ideTotal = CompLineCount(name)
    root("ideTotalLines") = ideTotal
    If ideTotal <> LineCount(paneNow) Then root("syncWarning") = SyncWarningText(ideTotal, LineCount(paneNow))
    Set root("coords") = CoordsDictOfName(name, 1, "ide")
    ' TASK-P1-059: segmented timing (was ok/ideTotalLines only).
    root("writeCalls") = wCalls
    root("hunks") = m_lastWriteHunks
    root("readMs") = tRead - tEntry
    root("predictMs") = tPredict - tRead
    root("lintMs") = tLint - tPredict
    root("writeMs") = tWrite - tLint
    AddWritePhases root
    root("totalMs") = GetTickCount() - tEntry
    CopyLintPack root, pack
    BreakpointEdit = ReceiptJson(root)
End Function

' TASK-P1-073 / F2-A: the catalogue is now HONEST about how each entry can actually be
' driven. executableVia = control_execution (a real control_execution mode) |
' native_breakpoint (the Debug-menu toggle, no source marker) | command_bars (only the
' CommandBars execution path reaches it) | command_bars:open_only (the entry only OPENS a
' debug window - no tool can READ that window's contents). viaNote carries the reason when
' the entry is NOT reachable through control_execution. Pure OUTPUT addition: no parameter
' and no tool description depends on it, so no re-freeze is involved.
Private Sub AddDbg(ByVal arr As Object, ByVal id As String, ByVal keys As String, ByVal english As String, ByVal via As String, Optional ByVal viaNote As String = "")
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("id") = id
    d("keys") = keys
    d("english") = english
    d("executableVia") = via
    If Len(viaNote) > 0 Then d("viaNote") = viaNote
    ' Card TASK-P1-007 (authorized scope): expose the LIVE CommandBars Id and
    ' shortcutText next to the static catalogue, so a caller can drive the menu
    ' without any localized text. Resolution reuses the Id+shortcut rule, so an
    ' ambiguous Id is never trusted alone; when the control cannot be resolved
    ' live we report the values measured in TASK-P1-006 and say so.
    Dim sc As String
    Dim wantId As Long
    sc = ""
    wantId = CmdId(id, sc)
    Dim c As Object
    Set c = FindBarControlById(m_vbe, "", english, wantId, sc)
    On Error Resume Next
    If Not c Is Nothing Then
        d("menuId") = c.Id
        d("shortcutText") = c.ShortcutText
        d("caption") = c.Caption
        d("idSource") = "live"
    Else
        d("menuId") = wantId
        d("shortcutText") = sc
        d("idSource") = "table"
    End If
    On Error GoTo 0
    arr.Add d
End Sub

Private Function ToolDebugCommands() As String
    Dim arr As Object, root As Object
    Set arr = New Collection
    AddDbg arr, "run", "F5", "Start", "control_execution"
    AddDbg arr, "break", "Ctrl+Break", "Break", "control_execution"
    AddDbg arr, "end", "Shift+F5", "End", "control_execution"
    AddDbg arr, "step_into", "F8", "Step Into", "control_execution"
    AddDbg arr, "step_over", "Shift+F8", "Step Over", "control_execution"
    AddDbg arr, "step_out", "Ctrl+Shift+F8", "Step Out", "control_execution"
    AddDbg arr, "run_to_cursor", "Ctrl+F8", "Run To Cursor", "control_execution"
    ' TASK-P1-073 / F2 (S ruling: STOP DECLARING it as a control_execution mode; F2-B is NOT
    ' implemented). The entry stays because the command DOES exist in the IDE and IS reachable
    ' - but only through the CommandBars path (command_bars{action:"execute"} by id/caption),
    ' never through control_execution, whose Select Case has no branch for it.
    AddDbg arr, "set_next_statement", "Ctrl+F9", "Set Next Statement", "command_bars", "control_execution has no branch for this mode"
    AddDbg arr, "toggle_breakpoint", "F9", "Toggle Breakpoint", "native_breakpoint"
    AddDbg arr, "clear_all_breakpoints", "Ctrl+Shift+F9", "Clear All Breakpoints", "native_breakpoint"
    AddDbg arr, "immediate_window", "Ctrl+G", "Immediate Window", "command_bars:open_only"
    AddDbg arr, "locals_window", "", "Locals Window", "command_bars:open_only"
    AddDbg arr, "watch_window", "", "Watch Window", "command_bars:open_only"
    Set root = CreateObject("Scripting.Dictionary")
    Set root("commands") = arr
    root("note") = "Fixed catalog. Each entry carries executableVia: control_execution (a real control_execution mode), native_breakpoint, or command_bars:open_only (the entry only OPENS a debug window - no tool can READ its contents). Captions are localized (the Chinese IDE uses different labels)."
    ToolDebugCommands = JsonStringify(root)
End Function

Private Function ToolGenerateStub(ByVal name As String) As String
    ToolGenerateStub = StubGen.StubJson(name, CompText(name))
End Function

Private Function ToolDependencyGraph() As String
    ' TASK-P1-058: report the wall time of the whole build (the 179 s could not be
    ' attributed without it) and the total characters scanned, so a regression is
    ' visible in the receipt instead of as an unresponsive IDE.
    Dim t0 As Long
    t0 = GetTickCount()
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    Dim comp As VBIDE.VBComponent
    Dim cm As VBIDE.CodeModule
    Dim chars As Double
    DepGraph.DepReset
    On Error Resume Next
    For Each comp In p.VBComponents
        MaybePump
        Err.Clear
        Set cm = comp.CodeModule
        If Not cm Is Nothing Then
            If cm.CountOfLines > 0 Then
                Dim tx As String
                tx = cm.Lines(1, cm.CountOfLines)
                chars = chars + Len(tx)
                DepGraph.DepAdd comp.Name, tx
            End If
        End If
    Next
    On Error GoTo 0
    Dim js As String
    js = DepGraph.DepGraphJson()
    ' additive: splice ms/chars straight after the opening brace
    ToolDependencyGraph = "{" & Chr$(34) & "ms" & Chr$(34) & ":" & CStr(GetTickCount() - t0) _
        & "," & Chr$(34) & "chars" & Chr$(34) & ":" & CStr(CLng(chars)) & "," & Mid$(js, 2)
End Function

Private Function ToolPurityCheck(ByVal name As String) As String
    DepGraph.DepReset
    DepGraph.DepAdd name, CompText(name)
    ToolPurityCheck = DepGraph.DepGraphJson()
End Function

Private Function ToolExportPublicApi(ByVal name As String) As String
    ToolExportPublicApi = ApiSig.ApiExportJson(CompText(name))
End Function

Private Function ToolCheckPublicApi(ByVal name As String, ByVal baseline As String) As String
    If Len(baseline) = 0 Then Err.Raise vbObjectError + 95, "CodeTools", "baseline spec is required"
    ToolCheckPublicApi = ApiSig.ApiDiffSpecJson(baseline, CompText(name))
End Function

Private Function ToolExecutionState(ByVal pid As Long) As String
    If pid <= 0 Then pid = WinProbe.CurrentProcessId()
    ToolExecutionState = WinProbe.ExecutionStateJson(pid)
End Function

' TASK-P1-024 helpers ---------------------------------------------------------

' gamma-4: validate="strict" refuses the write when the PRE-WRITE lint found
' severity=error diagnostics; "off" skips lint; "warn" (default) keeps the A-batch
' behaviour. Returns "" when the write may proceed, otherwise the refusal JSON.
Private Function StrictRefusal(ByVal target As String, ByVal toolName As String, ByVal pack As Object, ByVal validateMode As String) As String
    StrictRefusal = ""
    If LCase$(Trim$(validateMode)) <> "strict" Then Exit Function
    If pack Is Nothing Then Exit Function
    Dim es As Long
    es = 0
    On Error Resume Next
    es = CLng(pack("lintSummary")("errors"))
    On Error GoTo 0
    If es <= 0 Then Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "lint_errors"
    d("component") = target
    d("tool") = toolName
    d("validate") = "strict"
    d("lintErrors") = es
    d("detail") = "validate:'strict' refused the write: the result has " & CStr(es) & " severity=error diagnostic(s). Nothing was written (IDE and file are unchanged)."
    CopyLintPack d, pack
    StrictRefusal = JsonStringify(d)
End Function

' ===========================================================================
' TASK-P1-110: structured refusals for a bad ENUM argument and for a missing
' project file. Both are caller-input conditions, so both answer with the same
' verdict vocabulary the P1-108 classifier uses (reason / argument / allowedValues
' / suggestedAction / sideEffects) instead of silently accepting the value.
' ===========================================================================
Private Function EnumArgRefusalJson(ByVal tool As String, ByVal argName As String, ByVal badValue As String, ByVal reason As String, ByVal allowed As Object) As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("tool") = tool
    d("argument") = argName
    d("value") = badValue
    Set d("allowedValues") = allowed
    Dim near As String
    near = NearMatchOf(badValue, allowed)
    If Len(near) > 0 Then d("nearMatch") = near
    Dim sug As String
    If reason = "required_argument_missing" Then
        d("message") = argName & " is required (expected " & JoinList(allowed, "/") & ")"
        sug = "supply the reported argument (see 'argument' / 'allowedValues') and retry"
    Else
        d("message") = "Invalid " & argName & ": " & badValue & " (expected " & JoinList(allowed, "/") & ")"
        sug = "use one of allowedValues and retry"
        If Len(near) > 0 Then sug = sug & " (did you mean '" & near & "'?)"
    End If
    d("suggestedAction") = sug
    d("sideEffects") = "none: the call was refused BEFORE any IDE access or write"
    EnumArgRefusalJson = JsonStringify(d)
End Function

' The nearest legal value at Levenshtein distance <= 2, or "" when nothing is close.
Private Function NearMatchOf(ByVal bad As String, ByVal allowed As Object) As String
    On Error Resume Next
    Dim v As Variant, best As String, bestD As Long, dd As Long
    best = ""
    bestD = 3
    For Each v In allowed
        dd = Levenshtein(LCase$(Trim$(bad)), LCase$(CStr(v)))
        If dd < bestD Then
            bestD = dd
            best = CStr(v)
        End If
    Next
    NearMatchOf = best
End Function

Private Function Levenshtein(ByVal a As String, ByVal b As String) As Long
    Dim la As Long, lb As Long, i As Long, j As Long, cost As Long
    la = Len(a)
    lb = Len(b)
    If la = 0 Then
        Levenshtein = lb
        Exit Function
    End If
    If lb = 0 Then
        Levenshtein = la
        Exit Function
    End If
    Dim prev() As Long, cur() As Long
    ReDim prev(0 To lb)
    ReDim cur(0 To lb)
    For j = 0 To lb
        prev(j) = j
    Next j
    For i = 1 To la
        cur(0) = i
        For j = 1 To lb
            cost = 1
            If Mid$(a, i, 1) = Mid$(b, j, 1) Then cost = 0
            cur(j) = prev(j) + 1
            If cur(j - 1) + 1 < cur(j) Then cur(j) = cur(j - 1) + 1
            If prev(j - 1) + cost < cur(j) Then cur(j) = prev(j - 1) + cost
        Next j
        For j = 0 To lb
            prev(j) = cur(j)
        Next j
    Next i
    Levenshtein = prev(lb)
End Function

Private Function JoinList(ByVal c As Object, ByVal sep As String) As String
    Dim v As Variant, s As String
    For Each v In c
        If Len(s) > 0 Then s = s & sep
        s = s & CStr(v)
    Next
    JoinList = s
End Function

Private Function PropValidate() As Object
    Dim p As Object
    Set p = Prop("string", "off = no lint, warn (default) = report diagnostics and still write, strict = refuse the write when a severity=error diagnostic exists")
    Set p("enum") = EnumList("off", "warn", "strict")
    Set PropValidate = p
End Function

' gamma-5/6/7: resolve one anchor (+ occurrence/match/scope) to a concrete IDE line.
' requireUnique keeps the A-batch semantics (an ambiguous anchor is refused) when the
' caller did not ask for a specific occurrence; giving occurrence:N allows repeats.
' TASK-P1-097 convergence: the batch index is supplied by the caller (iRaw/iNorm/iFirst/
' iCnts, built ONCE per batch with VbAnchor.BuildNormIndex) so a unique anchor is answered
' in O(1) and matchedText comes off iRaw instead of a per-op TextLines. The candidate
' HitList - a whole-text split PLUS one CodeModule.ProcOfLine COM call per hit - is no
' longer built here: the caller builds it only when it has to REPORT a failure, so the
' success path of a batch makes ZERO ProcOfLine calls. With no index (caller passed
' nothing, e.g. the single-op edit_by_anchor path) the old exact linear scan is used, so
' those receipts cannot move.
Private Function ResolveAnchor(ByVal name As String, ByVal text As String, ByVal anchor As String, _
                               ByVal matchMode As String, ByVal scopeObj As Object, _
                               ByVal occurrence As Long, ByVal requireUnique As Boolean, _
                               ByRef lineOut As Long, ByRef matchedText As String, _
                               ByRef reason As String, ByRef hits() As Long, ByRef hitCount As Long, _
                               ByRef spanLines As Long, _
                               ByRef iRaw() As String, ByRef iNorm() As String, _
                               ByRef iFirst As Object, ByRef iCnts As Object, _
                               ByRef iRawFirst As Object, ByRef iRawCnts As Object) As Boolean
    ResolveAnchor = False
    lineOut = 0
    matchedText = ""
    reason = ""
    spanLines = 1
    Dim scopeStart As Long, scopeEnd As Long
    scopeStart = 0
    scopeEnd = 0
    If Not scopeObj Is Nothing Then
        scopeStart = JsonGetLong(scopeObj, "start_line", 0)
        scopeEnd = JsonGetLong(scopeObj, "end_line", 0)
        Dim sp As String
        sp = DictStr(scopeObj, "proc")
        If Len(sp) > 0 Then
            Dim cm As VBIDE.CodeModule
            On Error Resume Next
            Set cm = FindComponent(name).CodeModule
            On Error GoTo 0
            If cm Is Nothing Then
                reason = "scope_proc_not_found"
                Exit Function
            End If
            Dim ps As Long, pe As Long
            If Not ProcRangeOf(cm, sp, ps, pe) Then
                reason = "scope_proc_not_found"
                Exit Function
            End If
            scopeStart = ps
            scopeEnd = pe
        End If
    End If
    Dim cnt As Long
    ' TASK-P1-126 (MCP-07): a MULTI-LINE anchor matches a BLOCK - every anchor line
    ' consecutively - which is exactly what a caller who pasted three lines expected.
    ' A single-line anchor keeps the old path (index when the batch built one, else the
    ' linear scan), so those receipts cannot move.
    Dim blkN As Long
    blkN = 1
    If InStr(1, anchor, vbLf) > 0 Or InStr(1, anchor, vbCr) > 0 Then blkN = VbAnchor.AnchorLineCount(anchor)
    If blkN > 1 Then
        blkN = VbAnchor.FindBlockHitsEx(text, anchor, matchMode, scopeStart, scopeEnd, hits, cnt)
        If blkN < 1 Then blkN = 1
        spanLines = blkN
    Else
        If LineUbound(iNorm) >= 0 Then
            FindHitsIdx text, iRaw, iNorm, iFirst, iCnts, iRawFirst, iRawCnts, anchor, matchMode, scopeStart, scopeEnd, hits, cnt
        Else
            VbAnchor.FindHitsEx text, anchor, matchMode, scopeStart, scopeEnd, hits, cnt
        End If
    End If
    hitCount = cnt
    If cnt = 0 Then
        reason = "anchor_not_found"
        Exit Function
    End If
    Dim idx As Long
    If occurrence > 0 Then
        If occurrence > cnt Then
            reason = "anchor_not_found"
            Exit Function
        End If
        idx = occurrence - 1
    Else
        If requireUnique And cnt > 1 Then
            reason = "anchor_not_unique"
            Exit Function
        End If
        idx = 0
    End If
    lineOut = hits(idx)
    Dim mSrc() As String
    If LineUbound(iRaw) >= 0 Then
        mSrc = iRaw
    Else
        mSrc = TextLines(text)
    End If
    If lineOut >= 1 And lineOut <= UBound(mSrc) + 1 Then
        If spanLines > 1 Then
            ' TASK-P1-126 (MCP-07): the matched TEXT of a block anchor is the whole
            ' block, joined with vbCrLf - the receipt must show what it really matched.
            Dim bEnd As Long, bi As Long
            bEnd = lineOut + spanLines - 1
            If bEnd > UBound(mSrc) + 1 Then bEnd = UBound(mSrc) + 1
            matchedText = ""
            For bi = lineOut To bEnd
                If Len(matchedText) > 0 Then matchedText = matchedText & vbCrLf
                matchedText = matchedText & mSrc(bi - 1)
            Next bi
        Else
            matchedText = mSrc(lineOut - 1)
        End If
    End If
    ResolveAnchor = True
End Function

' TASK-P1-097: UBound of a Split result, safely -1 for an empty array (same helper shape
' as EditTxn's, which is module-private there and therefore not reachable from here).
Private Function LineUbound(ByRef arr() As String) As Long
    Dim u As Long
    u = -1
    On Error Resume Next
    u = UBound(arr)
    On Error GoTo 0
    LineUbound = u
End Function

' TASK-P1-097: hits for ONE anchor against the batch index. Byte-identical to
' VbAnchor.FindHitsEx (same order, same array shape, the same empty-pane convention), but
' O(1) when the anchor is UNIQUE and O(scope) when a scope is given. An ambiguous anchor
' still takes the exact linear scan, so its hit list and occurrence=k semantics cannot
' move. VbAnchor.CountLines is deliberately NOT used here: the index already knows the
' line count, and the old code called CountLines twice per op on the whole text.
Private Sub FindHitsIdx(ByVal text As String, ByRef iRaw() As String, ByRef iNorm() As String, _
                        ByRef iFirst As Object, ByRef iCnts As Object, _
                        ByRef iRawFirst As Object, ByRef iRawCnts As Object, _
                        ByVal anchor As String, ByVal matchMode As String, _
                        ByVal scopeStart As Long, ByVal scopeEnd As Long, _
                        ByRef hits() As Long, ByRef cnt As Long)
    Dim exact As Boolean, i As Long, lo As Long, hi As Long, k As String, nk As String
    exact = (LCase$(Trim$(matchMode)) = "exact")
    If exact Then EnsureRawIndex iRaw, iRawFirst, iRawCnts
    ' TASK-P1-063 / FindHitsEx: an EMPTY pane has NO hits (Split("") is an empty array).
    If Len(text) = 0 Then
        ReDim hits(0 To 0)
        cnt = 0
        Exit Sub
    End If
    nk = VbAnchor.NormLine(anchor)
    If scopeStart > 0 Or scopeEnd > 0 Then
        ' Same filter as FindHitsEx: out-of-scope lines are skipped, hits stay ascending.
        lo = 1
        If scopeStart > 0 Then lo = scopeStart
        hi = LineUbound(iRaw) + 1
        If scopeEnd > 0 And scopeEnd < hi Then hi = scopeEnd
        ReDim hits(0 To LineUbound(iRaw))
        cnt = 0
        For i = lo To hi
            If exact Then
                If StrComp(iRaw(i - 1), anchor, vbBinaryCompare) = 0 Then
                    hits(cnt) = i
                    cnt = cnt + 1
                End If
            ElseIf iNorm(i - 1) = nk Then
                hits(cnt) = i
                cnt = cnt + 1
            End If
        Next
        Exit Sub
    End If
    If exact Then
        k = anchor
        If Not iRawCnts.Exists(k) Then
            ReDim hits(0 To 0)
            cnt = 0
            Exit Sub
        End If
        cnt = CLng(iRawCnts(k))
        If cnt = 1 Then
            ReDim hits(0 To 0)
            hits(0) = CLng(iRawFirst(k))
            Exit Sub
        End If
    Else
        k = nk
        If Not iCnts.Exists(k) Then
            ReDim hits(0 To 0)
            cnt = 0
            Exit Sub
        End If
        cnt = CLng(iCnts(k))
        If cnt = 1 Then
            ReDim hits(0 To 0)
            hits(0) = CLng(iFirst(k))
            Exit Sub
        End If
    End If
    ' Ambiguous anchor: exact linear scan, so the hit LIST and its order are identical to
    ' FindHitsEx (callers report "anchor not unique" plus every candidate line).
    ReDim hits(0 To LineUbound(iRaw))
    cnt = 0
    If exact Then
        For i = 0 To LineUbound(iRaw)
            If StrComp(iRaw(i), anchor, vbBinaryCompare) = 0 Then
                hits(cnt) = i + 1
                cnt = cnt + 1
            End If
        Next
    Else
        For i = 0 To LineUbound(iRaw)
            If iNorm(i) = k Then
                hits(cnt) = i + 1
                cnt = cnt + 1
            End If
        Next
    End If
End Sub

' TASK-P1-097: the raw-line index is built LAZILY - only a batch that uses
' match:"exact" pays for it.
Private Sub EnsureRawIndex(ByRef iRaw() As String, ByRef rawFirst As Object, ByRef rawCnts As Object)
    If Not (rawFirst Is Nothing) Then Exit Sub
    Set rawFirst = CreateObject("Scripting.Dictionary")
    Set rawCnts = CreateObject("Scripting.Dictionary")
    Dim i As Long, k As String
    For i = 0 To LineUbound(iRaw)
        MaybePump
        k = iRaw(i)
        If rawFirst.Exists(k) Then
            rawCnts(k) = CLng(rawCnts(k)) + 1
        Else
            rawFirst(k) = i + 1
            rawCnts(k) = 1
        End If
    Next
End Sub

Private Function ScopeObjOf(ByVal args As Object, ByVal key As String) As Object
    On Error Resume Next
    Set ScopeObjOf = Nothing
    Set ScopeObjOf = args(key)
End Function

' ================= TASK-P1-024 delta tools (read-only) ======================

' delta-1: unsaved components in the IDE next to uncommitted files in git.
' READ-ONLY: it only reads the component model and runs read-only git subcommands.
Private Function ToolWorkspaceStatus(ByVal repoArg As String) As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim dirty As Object
    Set dirty = New Collection
    Dim prj As VBIDE.VBProject
    Dim openCount As Long
    openCount = 0
    On Error Resume Next
    Set prj = GetProject()
    On Error GoTo 0
    Dim i As Long
    If Not prj Is Nothing Then
        On Error Resume Next
        openCount = prj.VBComponents.Count
        On Error GoTo 0
        For i = 1 To openCount
            Dim comp As VBIDE.VBComponent
            On Error Resume Next
            Set comp = prj.VBComponents(i)
            On Error GoTo 0
            If Not comp Is Nothing Then
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("name") = comp.Name
                d("file") = ComponentFileNameOf(comp)
                Dim dr As Boolean
                dr = False
                On Error Resume Next
                dr = comp.IsDirty
                On Error GoTo 0
                d("isDirty") = dr
                If dr Then dirty.Add d
            End If
        Next
    End If
    Set root("dirtyComponents") = dirty
    root("dirtyCount") = dirty.Count
    root("ideOpenComponents") = openCount
    root("readOnly") = True
    Dim repo As String
    repo = repoArg
    If Len(repo) = 0 Then repo = ActiveProjectDir()
    Dim git As Object
    Set git = CreateObject("Scripting.Dictionary")
    If Len(repo) = 0 Then
        git("available") = False
        git("note") = "no project is open in the IDE and no repo argument was given"
    Else
        repo = ResolvePath(repo)
        Dim st As String, ds As String, so As String
        st = ""
        ds = ""
        so = ""
        ' WScript.Shell.Exec only drains the pipe after the child exits, so a git
        ' status whose output exceeds the pipe buffer DEADLOCKS and looks like a
        ' timeout (measured). Redirect to a temp file and read the file instead.
        Dim tmpF As String
        tmpF = Environ$("TEMP") & "\vb6mcp_git_status.txt"
        On Error Resume Next
        st = ProjShip.RunCaptureJson("cmd.exe", "/c git -C " & Qq(repo) & " status --porcelain=v1 -b > " & Qq(tmpF) & " 2>&1", 60000)
        so = ""
        so = Build.ReadLogAnsi(tmpF)
        ds = ProjShip.RunCaptureJson("git.exe", "-C " & Qq(repo) & " --no-pager diff --stat", 30000)
        On Error GoTo 0
        git("repo") = repo
        git("available") = (Len(so) >= 0)
        git("branch") = BranchOfPorcelain(so)
        git("ahead") = CountTag(so, "ahead ")
        git("behind") = CountTag(so, "behind ")
        Dim fls As Object
        Set fls = ParsePorcelain(so)
        Set git("files") = fls
        git("fileCount") = fls.Count
        git("statusText") = so
        git("diffStat") = JsonUnescape(JsonFieldRaw(ds, "stdout"))
        git("exitCode") = CLng(Val("0" & JsonFieldRaw(st, "exitCode")))
        If Len(JsonFieldRaw(st, "error")) > 0 Then git("error") = JsonUnescape(JsonFieldRaw(st, "error"))
        If JsonFieldRaw(st, "timedOut") = "true" Then git("timedOut") = True
    End If
    Set root("git") = git
    ToolWorkspaceStatus = JsonStringify(root)
End Function

' delta-2: IDE-vs-disk + lint for a set of components, summary first and only the
' FAILING components in detail. READ-ONLY.
Private Function ToolVerifyScope(ByVal comps As Object, ByVal pathArg As String) As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    ' TASK-P1-061: segmented timing (enumerate / read / lint / compare / nested verify).
    Dim t0 As Long, tEnum As Long, tSeg As Long, tRead As Long, tLint As Long, tCmp As Long, tVer As Long
    t0 = GetTickCount()
    Dim prj As VBIDE.VBProject
    On Error Resume Next
    Set prj = GetProject()
    On Error GoTo 0
    Dim names As Object
    Set names = New Collection
    Dim it As Variant
    If Not comps Is Nothing Then
        For Each it In comps
            MaybePump
            names.Add CStr(it)
        Next
    ElseIf Not prj Is Nothing Then
        Dim i As Long
        For i = 1 To prj.VBComponents.Count
            MaybePump
            names.Add prj.VBComponents(i).Name
        Next
    End If
    tEnum = GetTickCount()
    Dim passN As Long, failN As Long, checkedN As Long
    ' TASK-P1-121 item 9 (F-a): OVERLAY counters. passN/failN are untouched, so the
    ' identity pass + fail == checked still holds; unknownN/unknowns only LABEL the
    ' subset of the fails that were merely UNPROVEN (verify answered UNKNOWN) rather than
    ' proven bad (FAIL).
    Dim unknownN As Long
    Dim unknowns As Object
    Set unknowns = New Collection
    Dim fails As Object
    Set fails = New Collection
    Dim nm As Variant
    For Each nm In names
        MaybePump
        Dim cur As String, dp As String, pack As Object
        tSeg = GetTickCount()
        cur = CompText(CStr(nm))
        tRead = tRead + (GetTickCount() - tSeg)
        tSeg = GetTickCount()
        dp = DiskPathOf(CStr(nm), pathArg)
        ' TASK-P1-061: the LOCAL whole-module lint (LintPack) cost ~1.1 s on the 27,000-line
        ' fixture, and the nested ToolVerify below lints the SAME text again - so the same
        ' module was linted TWICE. The nested verify is the authoritative verdict, so its
        ' lint counts are reused (same text, same engine, measured identical). This removes
        ' duplicate work, not a check.
        tLint = tLint + (GetTickCount() - tSeg)
        Dim ideEq As Boolean
        ideEq = False
        tSeg = GetTickCount()
        On Error Resume Next
        If Len(dp) > 0 Then
            If Dir$(dp) <> "" Then ideEq = DiskMatchesPane(CStr(nm), cur)
        End If
        On Error GoTo 0
        tCmp = tCmp + (GetTickCount() - tSeg)
        ' The scope verdict must agree with calling verify on each component, so run the
        ' authoritative per-component verify ONCE and reuse it for BOTH the verdict and the
        ' lint counts. TASK-P1-061: the previous code linted the same text twice (its own
        ' LintPack + the lint inside this verify); on the 27,000-line fixture that duplicate
        ' was ~1.1 s of verify_scope's time.
        Dim vo As String, overall As String
        vo = ""
        tSeg = GetTickCount()
        On Error Resume Next
        vo = ToolVerify(CStr(nm), pathArg, "")
        On Error GoTo 0
        tVer = tVer + (GetTickCount() - tSeg)
        overall = JsonUnescape(JsonFieldRaw(vo, "overall"))
        If Len(overall) = 0 Then overall = "UNKNOWN"
        ' DiagnosticsJson emits "line" from m_diag(i).physLine - exactly what DiagLine
        ' returned as "editLine" - so these values equal the old FirstDiagText(pack) output.
        Dim lerr As String, lwarn As String, rid As String, rln As String, rmsg As String
        lerr = JsonFieldRaw(vo, "lintErrors")
        lwarn = JsonFieldRaw(vo, "lintWarnings")
        rid = JsonUnescape(JsonFieldRaw(vo, "ruleId"))
        rln = JsonFieldRaw(vo, "line")
        rmsg = JsonUnescape(JsonFieldRaw(vo, "message"))
        Dim lintErrN As Long, lintWarnN As Long, firstDiag As String
        lintErrN = 0
        If IsNumeric(lerr) Then lintErrN = CLng(Val(lerr))
        lintWarnN = 0
        If IsNumeric(lwarn) Then lintWarnN = CLng(Val(lwarn))
        firstDiag = ""
        If Len(rid) > 0 Then firstDiag = rid & " @line " & rln & ": " & rmsg
        ' The key ORDER of the detail object is unchanged from before TASK-P1-061.
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("component") = CStr(nm)
        d("file") = dp
        d("ideEqualsDisk") = ideEq
        d("lintErrors") = lintErrN
        d("lintWarnings") = lintWarnN
        d("firstDiagnostic") = firstDiag
        d("verifyOverall") = overall
        checkedN = checkedN + 1
        If overall = "PASS" Then
            passN = passN + 1
        Else
            ' UNKNOWN still counts as a fail (unchanged); `bucket` only says WHICH KIND
            ' of fail it was.
            failN = failN + 1
            If overall = "FAIL" Then
                d("bucket") = "fail"
            Else
                d("bucket") = "unknown"
                unknownN = unknownN + 1
                unknowns.Add d
            End If
            fails.Add d
        End If
    Next
    root("checked") = checkedN
    root("pass") = passN
    root("fail") = failN
    root("unknown") = unknownN
    Set root("unknowns") = unknowns
    root("totalMs") = GetTickCount() - t0
    root("enumerateMs") = tEnum - t0
    root("readCompTextMs") = tRead
    root("lintMs") = tLint
    root("compareMs") = tCmp
    root("verifyMs") = tVer
    Set root("failures") = fails
    root("readOnly") = True
    root("note") = "read-only sweep: each component is compared IDE-vs-disk and linted; only failing components are listed in detail"
    ToolVerifyScope = ReceiptJson(root)
End Function

' delta-3: machine-readable todo list from the process bus + git. READ-ONLY.
Private Function ToolTodoStatus() As String
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim base As String
    base = FindBusRoot()
    If Len(base) = 0 Then base = ProjectDir()
    If Len(base) = 0 Then base = App.Path
    Dim tasks As Object, reps As Object
    Set tasks = New Collection
    Set reps = New Collection
    CollectMd base & "\_bus\tasks\P1", tasks
    CollectMd base & "\_bus\reports\P1", reps
    Dim openN As Long, i As Long
    openN = 0
    For i = 1 To tasks.Count
        MaybePump
        Dim st As String
        st = LCase$(CStr(tasks(i)("status")))
        If st <> "done" And st <> "verified" Then openN = openN + 1
    Next
    root("taskCount") = tasks.Count
    root("openTasks") = openN
    Set root("tasks") = tasks
    root("reportCount") = reps.Count
    Set root("reports") = reps
    root("readOnly") = True
    root("note") = "read-only: _bus/tasks/P1 and _bus/reports/P1 are scanned for status + title"
    ToolTodoStatus = JsonStringify(root)
End Function

' ---------------------------------------------------------------------------
' contract v1.3.0 mechanism 3: machine-readable version identity. READ-ONLY.
' build.sha256 is read from the DLL on disk at call time; breaking[] is the
' machine-readable twin of the initialize.instructions BREAKING block.
' ---------------------------------------------------------------------------
' =====================================================================
' contract v1.6.0 (TASK-P1-076)
'   replace_in_files - cross-component literal/regex replace over the OPEN project
'   undo_edit        - explicit single-step restore of a component's PANE text
' =====================================================================

' The undo store lives next to the project (<project>\_undo), gated by the SAME switch
' as every other project-directory write (TASK-P1-077 ProjectDirWritesKilled): with
' VB6MCP_ALLOW_PROJECTDIR_WRITES=0 no snapshot is written and undo_edit reports none.
' Snapshot files are UTF-8 (binary Put/Get through Utf8.bas), so non-ASCII code survives.
Private Function UndoDir() As String
    Dim d As String
    d = ProjectDir()
    If Len(d) = 0 Then Exit Function
    UndoDir = CtlJoinPath(d, "_undo")
End Function

Private Function UndoStamp() As String
    UndoStamp = Format$(Now, "yyyymmdd-hhnnss") & "-" & Format$(GetTickCount() And &HFFFF&, "00000")
End Function

Private Function SafeFileStem(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    r = ""
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_" Or ch = "-" Then
            r = r & ch
        Else
            r = r & "_"
        End If
    Next
    SafeFileStem = r
End Function

Private Sub WriteTextUtf8File(ByVal path As String, ByVal text As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open path For Binary As #f
    If Len(text) > 0 Then
        Dim b() As Byte
        b = Utf8.Utf8Encode(text)
        Put #f, , b
    End If
    Close #f
    Err.Clear
End Sub

Private Function ReadTextUtf8File(ByVal path As String, ByRef ok As Boolean) As String
    On Error Resume Next
    ok = False
    Dim f As Integer, n As Long
    f = FreeFile
    Open path For Binary As #f
    n = LOF(f)
    If n > 0 Then
        Dim b() As Byte
        ReDim b(0 To n - 1)
        Get #f, , b
        ReadTextUtf8File = Utf8.Utf8DecodeBytes(b)
    End If
    Close #f
    If Err.Number = 0 Then ok = True
    Err.Clear
End Function

' TASK-P1-087: read the opt-in "snapshot" argument. The declared JSON type is boolean, but
' the card / FAQ both speak of snapshot=on, and JsonGetBool only accepts true/1 - so this
' reads the RAW value and treats true / 1 / on / yes (case-insensitive) as ON. Anything
' else - including an absent key - is OFF, which is what keeps the default zero-cost.
Private Function SnapshotOptIn(ByVal args As Object) As Boolean
    SnapshotOptIn = False
    If args Is Nothing Then Exit Function
    On Error Resume Next
    If Not args.Exists("snapshot") Then Exit Function
    Dim v As Variant
    v = args("snapshot")
    If IsNull(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        SnapshotOptIn = CBool(v)
        Err.Clear
        Exit Function
    End If
    Dim s As String
    s = LCase$(Trim$(CStr(v)))
    SnapshotOptIn = (s = "true" Or s = "1" Or s = "on" Or s = "yes")
    Err.Clear
End Function

' TASK-P1-076 (P8): called by the single text-write funnel just before the pane changes.
' Never fatal - a failed snapshot must not break the edit, and undo_edit simply reports
' that no snapshot exists.
Private Sub UndoSnapshot(ByVal compName As String, ByVal before As String)
    On Error Resume Next
    If Len(compName) = 0 Then Exit Sub
    If ProjectDirWritesKilled() Then Exit Sub
    Dim d As String
    d = UndoDir()
    If Len(d) = 0 Then Exit Sub
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    If Not fso.FolderExists(d) Then fso.CreateFolder d
    ' TASK-P1-085: capture the path and record it in THIS call's project-directory write
    ' manifest, so the receipt can enumerate it (write failures stay silent by design, and
    ' a snapshot that did not land is NOT reported).
    Dim snapPath As String
    snapPath = CtlJoinPath(d, SafeFileStem(compName) & "." & UndoStamp() & ".txt")
    WriteTextUtf8File snapPath, before
    If CtlFileExists(snapPath) Then
        ManifestAdd m_dirWrites, snapPath
        ' TASK-P1-121 item 6: the same path, also recorded for the sideEffects block.
        SeAdd m_seUndo, snapPath
    End If
    Err.Clear
End Sub

' CompText itself can raise on a 0-line component (REQ-20260909-047/049); both new tools
' must treat "no text" as "no hits", never as a failure.
Private Function SafeCompText(ByVal name As String) As String
    On Error Resume Next
    SafeCompText = CompText(name)
    Err.Clear
End Function

Private Function CountOccur(ByVal hay As String, ByVal needle As String) As Long
    If Len(needle) = 0 Then Exit Function
    Dim n As Long, p As Long
    n = 0
    p = 1
    Do
        p = InStr(p, hay, needle, vbBinaryCompare)
        If p = 0 Then Exit Do
        n = n + 1
        p = p + Len(needle)
    Loop
    CountOccur = n
End Function

' CR 归一化（窗格文本没有行尾 CRLF，磁盘文件有）——等价 oracle 必须先归一化。
Private Function CrNorm(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    CrNorm = t
End Function

Private Function NormPathKey(ByVal p As String, ByVal fso As Object) As String
    Dim s As String
    s = Trim$(p)
    If Len(s) = 0 Then Exit Function
    On Error Resume Next
    If fso.FolderExists(s) Then s = fso.GetFolder(s).Path
    If fso.FileExists(s) Then s = fso.GetFile(s).Path
    Err.Clear
    Do While Len(s) > 3
        If Right$(s, 1) = "\" Then s = Left$(s, Len(s) - 1) Else Exit Do
    Loop
    NormPathKey = s
End Function

' expected_shas[] entry lookup: the caller may name a component OR its file path.
Private Function ExpectedShaFor(ByVal arr As Object, ByVal compName As String, ByVal filePath As String) As String
    On Error Resume Next
    ExpectedShaFor = ""
    If arr Is Nothing Then Exit Function
    Dim i As Long, e As Object, key As String
    For i = 1 To arr.Count
        Set e = arr.Item(i)
        If Not e Is Nothing Then
            key = ""
            key = CStr(e("file"))
            If Len(key) > 0 Then
                If StrComp(key, compName, vbTextCompare) = 0 Then
                    ExpectedShaFor = LCase$(CStr(e("sha256")))
                    Exit Function
                End If
                If Len(filePath) > 0 Then
                    If StrComp(key, filePath, vbTextCompare) = 0 Then
                        ExpectedShaFor = LCase$(CStr(e("sha256")))
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Private Function LintCountOf(ByVal pack As Object, ByVal key As String) As Long
    On Error Resume Next
    LintCountOf = 0
    If pack Is Nothing Then Exit Function
    Dim v As Variant
    v = pack(key)
    If IsNumeric(v) Then LintCountOf = CLng(v)
End Function

' ---------------------------------------------------------------------
' P7: replace_in_files
' ---------------------------------------------------------------------
Private Function ToolReplaceInFiles(ByVal path As String, ByVal query As String, ByVal replacement As String, _
        ByVal useRegex As Boolean, ByVal dryRun As Boolean, ByVal expectedShas As Object, _
        ByVal validateMode As String, ByVal maxFiles As Long) As String
    Dim tEntry As Long
    tEntry = GetTickCount()
    Dim proj As VBIDE.VBProject
    Set proj = GetProject()
    If proj Is Nothing Then
        Err.Raise vbObjectError + 120, "CodeTools", "MCPREASON:no_project|open a project first|there is no active project to replace in; nothing was written"
    End If
    If Len(Trim$(query)) = 0 Then
        Err.Raise vbObjectError + 121, "CodeTools", "MCPREASON:empty_query|pass a non-empty query|an empty query matches every position; nothing was written"
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pf As String
    pf = ""
    On Error Resume Next
    pf = proj.FileName
    Err.Clear
    On Error GoTo 0
    ' path (when given) must name THIS project: the tool never writes a project the
    ' add-in is not attached to, so "only ever a copy" is enforced, not just advised.
    If Len(Trim$(path)) > 0 Then
        Dim want As String, okPath As Boolean
        want = NormPathKey(path, fso)
        okPath = False
        If Len(pf) > 0 Then
            If StrComp(want, NormPathKey(pf, fso), vbTextCompare) = 0 Then okPath = True
            If StrComp(want, NormPathKey(fso.GetParentFolderName(pf), fso), vbTextCompare) = 0 Then okPath = True
        End If
        If Not okPath Then
            Err.Raise vbObjectError + 122, "CodeTools", "MCPREASON:not_the_open_project|pass the open project's .vbp or its directory|replace_in_files only rewrites the project this add-in is attached to (" & pf & "); nothing was written"
        End If
    End If
    If maxFiles <= 0 Then maxFiles = 200
    If maxFiles > 1000 Then maxFiles = 1000
    Dim re As Object
    If useRegex Then
        On Error Resume Next
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = False
        re.Pattern = query
        If Err.Number <> 0 Then
            Dim em As String
            em = Err.Description
            Err.Clear
            On Error GoTo 0
            Err.Raise vbObjectError + 123, "CodeTools", "MCPREASON:bad_regex|fix the pattern|the regex engine rejected the pattern: " & em & "; nothing was written"
        End If
        On Error GoTo 0
    End If
    Dim files As Object, applied As Object, appliedText As Object
    Set files = New Collection
    Set applied = New Collection
    Set appliedText = CreateObject("Scripting.Dictionary")
    Dim scanned As Long, changedFiles As Long, totalHits As Long, wouldChange As Long
    Dim stopped As Boolean, rolledBack As Boolean, stoppedAt As String, note As String
    Dim c As VBIDE.VBComponent
    For Each c In proj.VBComponents
        If scanned >= maxFiles Then
            note = "max_files cap reached (" & CStr(maxFiles) & "); the rest was NOT scanned"
            Exit For
        End If
        scanned = scanned + 1
        MaybePump
        Dim nm As String, fp As String, txt As String
        nm = c.Name
        fp = ComponentFileNameOf(c)
        Dim entry As Object
        Set entry = CreateObject("Scripting.Dictionary")
        entry("comp") = nm
        If Len(fp) > 0 Then entry("file") = fp
        entry("hits") = 0
        entry("changed") = False
        entry("skipped") = False
        entry("shaBefore") = ""
        entry("shaAfter") = ""
        txt = SafeCompText(nm)
        Dim cnt As Long, nt As String, firstHitLine As Long
        cnt = 0
        nt = ""
        firstHitLine = 0
        If Len(txt) > 0 Then
            If useRegex Then
                If re.Test(txt) Then
                    Dim ms As Object
                    Set ms = re.Execute(txt)
                    cnt = ms.Count
                    nt = re.Replace(txt, replacement)
                    Dim idx As Long
                    idx = ms.Item(0).FirstIndex
                    firstHitLine = CountOccur(Left$(txt, idx), vbCrLf) + 1
                End If
            Else
                cnt = CountOccur(txt, query)
                If cnt > 0 Then
                    nt = Replace(txt, query, replacement)
                    firstHitLine = CountOccur(Left$(txt, InStr(1, txt, query, vbBinaryCompare) - 1), vbCrLf) + 1
                End If
            End If
        End If
        If cnt = 0 Then
            files.Add entry
        Else
            totalHits = totalHits + cnt
            entry("hits") = cnt
            entry("firstHitLine") = firstHitLine
            entry("shaBefore") = Sha256.Sha256HexTextUtf8(txt)
            entry("shaAfter") = Sha256.Sha256HexTextUtf8(nt)
            Dim wantSha As String
            wantSha = ExpectedShaFor(expectedShas, nm, fp)
            If Len(wantSha) > 0 Then
                If StrComp(wantSha, CStr(entry("shaBefore")), vbTextCompare) <> 0 Then
                    entry("skipped") = True
                    entry("reason") = "expected_sha_mismatch"
                    entry("expectedSha256") = wantSha
                    files.Add entry
                    GoTo NextComp
                End If
            End If
            ' validate gate: same helper the other write tools use. strict => refuse to
            ' write that component (a real failure => stop + roll back).
            Dim pack As Object, sref As String
            Set pack = LintPackForEdit(nt, KindOfName(nm), LineCount(txt), firstHitLine)
            entry("lintErrors") = LintCountOf(pack, "lintErrors")
            entry("lintWarnings") = LintCountOf(pack, "lintWarnings")
            sref = StrictRefusal(nm, "replace_in_files", pack, validateMode)
            If Len(sref) > 0 Then
                entry("skipped") = True
                entry("reason") = "validate_strict_failed"
                If dryRun Then
                    files.Add entry
                    GoTo NextComp
                End If
                files.Add entry
                stopped = True
                stoppedAt = nm
                Exit For
            End If
            If dryRun Then
                wouldChange = wouldChange + 1
                files.Add entry
                GoTo NextComp
            End If
            ' APPLY: one minimal text transaction per component (per-component atomic)
            Dim wrote As Long
            wrote = SetCompTextMinimal(nm, txt, nt)
            If wrote <= 0 Then
                entry("reason") = "write_failed"
                files.Add entry
                stopped = True
                stoppedAt = nm
                Exit For
            End If
            Dim back As String
            back = SafeCompText(nm)
            If StrComp(CrNorm(back), CrNorm(nt), vbBinaryCompare) <> 0 Then
                entry("reason") = "pane_verify_failed"
                entry("shaAfter") = Sha256.Sha256HexTextUtf8(back)
                files.Add entry
                stopped = True
                stoppedAt = nm
                Exit For
            End If
            entry("changed") = True
            entry("shaAfter") = Sha256.Sha256HexTextUtf8(back)
            entry("writeCalls") = wrote
            entry("hunks") = m_lastWriteHunks
            changedFiles = changedFiles + 1
            applied.Add nm
            appliedText(nm) = txt
            files.Add entry
        End If
NextComp:
    Next c
    ' ---- honest failure semantics: stop + roll the applied components back ----
    If stopped Then
        rolledBack = True
        Dim ai As Long
        For ai = 1 To applied.Count
            Dim an As String, curTxt As String
            an = CStr(applied.Item(ai))
            curTxt = SafeCompText(an)
            If SetCompTextMinimal(an, curTxt, CStr(appliedText(an))) <= 0 Then rolledBack = False
            MaybePump
        Next
        If rolledBack Then changedFiles = 0
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = Not stopped
    ' TASK-P1-085: one snapshot per written component is enumerated here as well.
    Dim pdwR As Object
    Set pdwR = New Collection
    ManifestMergeUndo pdwR
    Set d("projectDirWrites") = pdwR
    If stopped Then
        d("reason") = "write_stopped"
        d("advice") = "the receipt's summary says exactly what was applied and whether the rollback completed; nothing is claimed to be atomic across components"
    End If
    d("dryRun") = dryRun
    d("regex") = useRegex
    Set d("files") = files
    Dim s As Object
    Set s = CreateObject("Scripting.Dictionary")
    s("files") = scanned
    s("filesChanged") = changedFiles
    s("wouldChange") = wouldChange
    s("totalHits") = totalHits
    Dim ap As Object
    Set ap = New Collection
    For ai = 1 To applied.Count
        ap.Add CStr(applied.Item(ai))
    Next
    Set s("applied") = ap
    If stopped Then
        s("stoppedAt") = stoppedAt
        s("rolledBack") = rolledBack
        If Not rolledBack Then
            s("note") = "STOPPED on " & stoppedAt & "; the rollback could NOT be completed - the applied set above is exactly what remains"
        End If
    End If
    If Len(note) > 0 Then
        If s.Exists("note") Then s("note") = s("note") & " | " & note Else s("note") = note
    End If
    Set d("summary") = s
    d("totalMs") = GetTickCount() - tEntry
    ToolReplaceInFiles = ReceiptJson(d)
End Function

' ---------------------------------------------------------------------
' P8: undo_edit
' ---------------------------------------------------------------------
' TASK-P1-087: the one sentence that turns "no snapshot" into an actionable next call.
Private Function NoSnapshotHint() As String
    NoSnapshotHint = "pass snapshot=true (""on""/""yes""/1 also accepted) to replace_code / apply_edits / edit_by_anchor to write a pre-edit pane snapshot; then undo_edit can restore it. set, apply_edits and edit_by_anchor already snapshot by default; replace_code's replace/insert/append/delete do NOT unless snapshot=true"
End Function

Private Function ToolUndoEdit(ByVal component As String, ByVal action As String, ByVal ts As String, ByVal toDisk As Boolean) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("action") = action
    ' TASK-P1-085: the manifest is the SAME Collection object for the whole call, so the
    ' snapshot written later by the restore path is visible when the receipt is serialised.
    Set d("projectDirWrites") = m_dirWrites
    If toDisk Then
        d("ok") = False
        d("reason") = "to_disk_not_supported"
        d("component") = component
        d("ts") = ts
        d("diskUntouched") = True
        d("advice") = "undo_edit restores the component's PANE text only; call save_project to flush the whole project to disk"
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim dir As String
    dir = UndoDir()
    d("undoDir") = dir
    If Len(dir) = 0 Then
        d("ok") = False
        d("reason") = "no_project_dir"
        d("advice") = "save the project first (save_project with a path); the undo store lives next to the project"
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then
        d("ok") = False
        ' TASK-P1-087: when the caller NAMED a component, the reason says so and points at the
        ' opt-in that creates one; the generic (no component) case keeps the old reason.
        If Len(component) > 0 Then
            d("reason") = "no_snapshot_for_component"
            d("component") = component
            d("hint") = NoSnapshotHint()
        Else
            d("reason") = "no_snapshot"
        End If
        ' TASK-P1-085: found live - assigning an object to a Dictionary item without Set
        ' raises VB error 450, so the "no snapshot store yet" path returned unexpected_error
        ' instead of the intended structured {ok:false, reason:"no_snapshot"}.
        Set d("snapshots") = New Collection
        d("advice") = "no snapshot exists yet: snapshots are written before every component text write (and are disabled while VB6MCP_ALLOW_PROJECTDIR_WRITES=0)"
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    ' collect all snapshots (newest ts LAST: the filename sorts by yyyymmdd-hhnnss-ttttt)
    Dim snaps As Object
    Set snaps = New Collection
    Dim f As Object
    For Each f In fso.GetFolder(dir).Files
        Dim leaf As String, p1 As Long, compPart As String, tsPart As String
        leaf = f.Name
        If LCase$(Right$(leaf, 4)) = ".txt" Then
            p1 = InStrRev(leaf, ".", Len(leaf) - 4)
            If p1 > 1 Then
                compPart = Left$(leaf, p1 - 1)
                tsPart = Mid$(leaf, p1 + 1, Len(leaf) - p1 - 4)
                Dim se As Object
                Set se = CreateObject("Scripting.Dictionary")
                se("component") = compPart
                se("file") = f.Path
                se("ts") = tsPart
                se("bytes") = f.Size
                snaps.Add se
            End If
        End If
    Next
    If LCase$(action) = "list" Then
        Dim sel As Object
        Set sel = New Collection
        Dim si As Long
        For si = 1 To snaps.Count
            If Len(component) = 0 Then
                sel.Add snaps.Item(si)
            ElseIf StrComp(CStr(snaps.Item(si)("component")), component, vbTextCompare) = 0 Then
                sel.Add snaps.Item(si)
            End If
        Next
        d("count") = sel.Count
        Set d("snapshots") = sel
        If sel.Count = 0 Then
            d("ok") = False
            ' TASK-P1-087: a NAMED component with no snapshot is its own structured case.
            If Len(component) > 0 Then
                d("reason") = "no_snapshot_for_component"
                d("component") = component
                d("hint") = NoSnapshotHint()
            Else
                d("reason") = "no_snapshot"
            End If
        End If
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    ' restore: the newest snapshot for that component (or the exact ts), unless the
    ' caller named a file. TS comparison is a prefix/equality on the stamp.
    Dim chosen As Object, ci As Long, cand As Object
    Set chosen = Nothing
    For ci = 1 To snaps.Count
        Set cand = snaps.Item(ci)
        If Len(component) = 0 Or StrComp(CStr(cand("component")), component, vbTextCompare) = 0 Then
            If Len(ts) = 0 Then
                Set chosen = cand
            ElseIf StrComp(CStr(cand("ts")), ts, vbTextCompare) = 0 Then
                Set chosen = cand
            End If
        End If
    Next
    If chosen Is Nothing Then
        d("ok") = False
        d("component") = component
        d("ts") = ts
        ' name what IS available so the caller can retry without guessing
        Dim av As Object
        Set av = New Collection
        For ci = 1 To snaps.Count
            If Len(component) = 0 Or StrComp(CStr(snaps.Item(ci)("component")), component, vbTextCompare) = 0 Then
                av.Add CStr(snaps.Item(ci)("ts"))
            End If
        Next
        Set d("available") = av
        ' TASK-P1-087: a named component with NO snapshot at all is a distinct structured case
        ' (and names the opt-in that creates one); a merely-absent ts on a component that DOES
        ' have snapshots keeps snapshot_not_found + available[].
        If Len(component) > 0 And av.Count = 0 Then
            d("reason") = "no_snapshot_for_component"
            d("hint") = NoSnapshotHint()
        Else
            d("reason") = "snapshot_not_found"
        End If
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim backOk As Boolean
    Dim snapText As String
    snapText = ReadTextUtf8File(CStr(chosen("file")), backOk)
    If Not backOk Then
        d("ok") = False
        d("reason") = "snapshot_unreadable"
        d("file") = chosen("file")
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim compName As String
    compName = CStr(chosen("component"))
    Dim comp As VBIDE.VBComponent
    Set comp = FindComponent(compName)
    If comp Is Nothing Then
        d("ok") = False
        d("reason") = "component_not_found"
        d("component") = compName
        d("advice") = "the component this snapshot belongs to is not in the open project; the snapshot file is left untouched"
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim cur As String
    cur = SafeCompText(compName)
    d("component") = compName
    d("ts") = chosen("ts")
    d("shaBefore") = Sha256.Sha256HexTextUtf8(cur)
    d("snapshotSha256") = Sha256.Sha256HexTextUtf8(snapText)
    If StrComp(CrNorm(cur), CrNorm(snapText), vbBinaryCompare) = 0 Then
        d("restored") = False
        d("shaAfter") = d("shaBefore")
        d("diskUntouched") = True
        d("note") = "the pane already equals that snapshot; nothing was written"
        ToolUndoEdit = JsonStringify(d)
        Exit Function
    End If
    Dim wr As Long
    wr = SetCompTextMinimal(compName, cur, snapText)
    d("writeCalls") = wr
    d("hunks") = m_lastWriteHunks
    Dim nowTxt As String
    nowTxt = SafeCompText(compName)
    d("shaAfter") = Sha256.Sha256HexTextUtf8(nowTxt)
    d("restored") = (StrComp(CrNorm(nowTxt), CrNorm(snapText), vbBinaryCompare) = 0)
    d("diskUntouched") = True
    If Not CBool(d("restored")) Then
        d("ok") = False
        d("reason") = "restore_verify_failed"
    End If
    ToolUndoEdit = JsonStringify(d)
End Function

' TASK-P1-088: attach the wheel snapshot and stamp the ONE pane-top channel that really
' answers on this IDE: VBIDE ActiveCodePane.TopLine (1-based, 0 = no active pane). The
' wheel object's own lastTopLineBefore/After come from EM_GETFIRSTVISIBLELINE and read -1
' because the code pane exposes no RichEdit20A child (measured); this key is the
' cross-checkable replacement. Pure output - no new tool, no description change.
Private Sub StampWheelNow(ByVal root As Object)
    Dim wo As Object
    Set wo = WinOps.WheelStateObj()
    wo("paneTopLineNow") = ActivePaneTopLine()
    Set root("wheel") = wo
End Sub

Private Function ActivePaneTopLine() As Long
    On Error Resume Next
    ActivePaneTopLine = 0
    Dim cp As VBIDE.CodePane
    Set cp = m_vbe.ActiveCodePane
    If cp Is Nothing Then Exit Function
    ActivePaneTopLine = cp.TopLine
    If Err.Number <> 0 Then ActivePaneTopLine = 0
    On Error GoTo 0
End Function

' =====================================================================
' TASK-P1-067 final fix: the optional code-pane wheel.
'
' The wheel subclasses the pane's own window (class "VbaWindow"), but the pane
' does NOT respond to WM_VSCROLL: measured 2026-10-09 on the real code pane
' (hwnd=23464986, "P1067 - modLong (Code)"), SB_LINEDOWN x2 / SB_PAGEDOWN /
' SB_LINEUP x2 all left CodePane.TopLine at 1 with SendMessage returning 0 -
' i.e. the message arrives and is ignored. Setting TopLine through the object
' model IS observable and reversible (4/4). WinOps.bas is a standard module with
' no VBIDE reference, so it reaches these three helpers through the late-bound
' sink that McpServer wires at start-up (WinOps.WheelSetSink m_tools).
'
' The mapping key is the window caption: measured GetWindowText(VbaWindow) ==
' CodePane.Window.Caption VERBATIM. The "(Code)" suffix is never parsed - a
' localised IDE changes it, and the whole string is what both sides carry.
' =====================================================================

' Captions of every code pane that is open right now. Used by WheelScan to tell a
' code pane apart from the Immediate / Watch / Locals windows, which share the
' same window class. Read-only.
Public Function WheelPaneCaptions() As Object
    On Error Resume Next
    Dim c As Object
    Set c = New Collection
    Dim cp As VBIDE.CodePane
    Dim n As Long
    n = 0
    For Each cp In m_vbe.CodePanes
        If Not cp Is Nothing Then
            If Not cp.Window Is Nothing Then
                c.Add cp.Window.Caption
                n = n + 1
            End If
        End If
    Next
    m_wheelCapsSource = "all"
    If n = 0 Then
        ' TASK-P1-132 (D-130-01), measured 2026-10-13: in BREAK mode the
        ' VBE.CodePanes collection yields NOTHING here - the pane the debugger opened
        ' exists, yet every VbaWindow was classified not_a_code_pane - while
        ' VBE.ActiveCodePane still answers (get_codepane{active} works in break).
        ' Fall back to the ACTIVE pane so a code pane that first appears IN BREAK can
        ' still be hooked. The source is published (WheelPaneCaptionsSource) so the
        ' receipt says why the table looks the way it does.
        Dim ap As VBIDE.CodePane
        Set ap = m_vbe.ActiveCodePane
        If Not ap Is Nothing Then
            If Not ap.Window Is Nothing Then
                c.Add ap.Window.Caption
                n = n + 1
                m_wheelCapsSource = "active_fallback"
            End If
        End If
        If n = 0 Then m_wheelCapsSource = "empty"
    End If
    Set WheelPaneCaptions = c
    Err.Clear
End Function

' TASK-P1-132 (D-130-01): read-only companion - WHERE the last caption table came from.
Public Function WheelPaneCaptionsSource() As String
    WheelPaneCaptionsSource = m_wheelCapsSource
End Function

' Current TopLine of the pane whose caption is caption; -1 when no pane matches.
Public Function WheelPaneTopByCaption(ByVal caption As String) As Long
    On Error Resume Next
    WheelPaneTopByCaption = -1
    m_wheelTopStatus = "ok"
    m_wheelTopErr = ""
    Dim cp As VBIDE.CodePane
    Set cp = FindCodePaneByCaption(caption)
    If cp Is Nothing Then
        m_wheelTopStatus = "pane_not_found"
        Exit Function
    End If
    Dim v As Long
    Err.Clear
    v = cp.TopLine
    If Err.Number <> 0 Then
        ' TASK-P1-132 (D-130-01): DO NOT collapse this into the same -1 as "not found".
        m_wheelTopStatus = "pane_not_writable_in_break"
        m_wheelTopErr = "TopLine READ refused: VB error " & CStr(Err.Number) & " (" & Err.Description & ")"
        Err.Clear
        Exit Function
    End If
    WheelPaneTopByCaption = v
    Err.Clear
End Function

' Move the viewport of the pane whose caption is caption, then READ IT BACK.
' wantLines / wantPages: exactly one is non-zero; the sign is the direction.
' Returns the pane's TopLine after the move, or -1 when no pane matches that
' caption (the caller must then NOT count the wheel as handled).
' Only TopLine is written - the caret and the selection are never touched.
Public Function WheelScrollPaneByCaption(ByVal caption As String, ByVal wantLines As Long, ByVal wantPages As Long) As Long
    On Error Resume Next
    WheelScrollPaneByCaption = -1
    m_wheelTopStatus = "ok"
    m_wheelTopErr = ""
    Dim cp As VBIDE.CodePane
    Set cp = FindCodePaneByCaption(caption)
    If cp Is Nothing Then
        m_wheelTopStatus = "pane_not_found"
        Exit Function
    End If
    Dim n As Long
    n = wantLines
    If wantPages <> 0 Then n = n + wantPages * cp.CountOfVisibleLines
    If n = 0 Then
        WheelScrollPaneByCaption = cp.TopLine
        Exit Function
    End If
    Dim lastLine As Long
    lastLine = 1
    If Not cp.CodeModule Is Nothing Then
        If cp.CodeModule.CountOfLines > 0 Then lastLine = cp.CodeModule.CountOfLines
    End If
    Dim t As Long
    t = cp.TopLine + n
    If t < 1 Then t = 1
    If t > lastLine Then t = lastLine
    Err.Clear
    cp.TopLine = t
    If Err.Number <> 0 Then
        ' TASK-P1-132 (D-130-01): the pane WAS resolved; the IDE refused the WRITE. That
        ' is a DIFFERENT failure from "no pane matched" and is reported as such, with the
        ' real VB error - never as the generic caption failure.
        m_wheelTopStatus = "pane_not_writable_in_break"
        m_wheelTopErr = "TopLine WRITE refused (target " & CStr(t) & " of " & CStr(lastLine) & "): VB error " & CStr(Err.Number) & " (" & Err.Description & ")"
        Err.Clear
        Exit Function
    End If
    ' read back what the IDE actually accepted, never what we hoped for
    Err.Clear
    WheelScrollPaneByCaption = cp.TopLine
    If Err.Number <> 0 Then
        m_wheelTopStatus = "pane_not_writable_in_break"
        m_wheelTopErr = "TopLine READ-BACK refused: VB error " & CStr(Err.Number) & " (" & Err.Description & ")"
        WheelScrollPaneByCaption = -1
    End If
    Err.Clear
End Function

' Verbatim (case-sensitive) caption match over every open code pane.
Private Function FindCodePaneByCaption(ByVal caption As String) As VBIDE.CodePane
    On Error Resume Next
    Set FindCodePaneByCaption = Nothing
    If Len(caption) = 0 Then Exit Function
    m_wheelResolveVia = "none"
    Dim cp As VBIDE.CodePane
    ' (1) verbatim Window.Caption match - the DESIGN-time channel, unchanged.
    For Each cp In m_vbe.CodePanes
        If Not cp Is Nothing Then
            If Not cp.Window Is Nothing Then
                If StrComp(cp.Window.Caption, caption, vbBinaryCompare) = 0 Then
                    Set FindCodePaneByCaption = cp
                    m_wheelResolveVia = "caption"
                    Exit Function
                End If
            End If
        End If
    Next
    ' (2) TASK-P1-132 (D-130-01), measured 2026-10-13: in BREAK mode CodePane.Window is
    ' not readable at all - get_codepane{active} answers TopLine/CodeModule but its
    ' windowCaption comes back EMPTY - so channel (1) can never match. Identify the pane
    ' by the COMPONENT NAME embedded in the window title ("<project> - <name> (Code)")
    ' and reach it through the SAME object path the active-pane tool uses
    ' (CodeModule.CodePane), which is readable in break.
    Dim nm As String
    nm = ComponentNameFromPaneCaption(caption)
    If Len(nm) > 0 Then
        Dim c2 As VBIDE.VBComponent
        Set c2 = FindComponent(nm)
        If Not c2 Is Nothing Then
            If Not c2.CodeModule Is Nothing Then
                Set FindCodePaneByCaption = c2.CodeModule.CodePane
                If Not FindCodePaneByCaption Is Nothing Then
                    m_wheelResolveVia = "component"
                    Exit Function
                End If
            End If
        End If
    End If
    ' (3) last resort: the ACTIVE pane, but ONLY when the caption really names it, so a
    ' non-code window (Immediate / Watch / Locals) can never be mapped onto it - that
    ' mapping is exactly the residual this task forbids.
    Dim ap As VBIDE.CodePane
    Set ap = m_vbe.ActiveCodePane
    If Not ap Is Nothing Then
        Dim an As String
        an = ""
        If Not ap.CodeModule Is Nothing Then an = ap.CodeModule.Name
        If Len(an) > 0 Then
            If InStr(1, caption, an, vbTextCompare) > 0 Then
                Set FindCodePaneByCaption = ap
                m_wheelResolveVia = "active_name"
            End If
        End If
    End If
    Err.Clear
End Function

' TASK-P1-132 (D-130-01): "<project> - <component> (Code)" -> "<component>".
' A trailing bracketed suffix is dropped whatever its language ("(Code)", "(代码)"),
' because the component name never contains a bracket here.
Private Function ComponentNameFromPaneCaption(ByVal caption As String) As String
    ComponentNameFromPaneCaption = ""
    Dim s As String
    s = caption
    Dim i As Long, q As Long
    q = 0
    For i = Len(s) To 1 Step -1
        If Mid$(s, i, 1) = "(" Or Mid$(s, i, 1) = "[" Then
            q = i
            Exit For
        End If
    Next
    If q > 0 Then s = RTrim$(Left$(s, q - 1))
    Dim p As Long
    p = InStrRev(s, " - ")
    If p > 0 Then s = Mid$(s, p + 3)
    ComponentNameFromPaneCaption = Trim$(s)
End Function

' TASK-P1-132 (D-130-01): read-only companion - which channel resolved the last lookup.
Public Function WheelResolveVia() As String
    WheelResolveVia = m_wheelResolveVia
End Function

' TASK-P1-132 (D-130-01): the REAL outcome of the last TopLine read/write, so a caller
' can tell "no pane matched" from "the pane was found but the IDE refused the write".
Public Function WheelTopStatus() As String
    WheelTopStatus = m_wheelTopStatus
End Function

Public Function WheelTopError() As String
    WheelTopError = m_wheelTopErr
End Function

Private Function ToolMcpInfo() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("product") = VB6MCP_PRODUCT_VERSION
    d("contract") = VB6MCP_CONTRACT_VERSION
    d("toolCount") = ToolCount()
    ' TASK-P1-039 (contract v1.3.2, optional fields): this build can decode a VB error code
    ' back to its HRESULT family via P2@s 85-entry table (R65/R66). The raw `hresult`
    ' field additionally needs dwMessageId@s RVA, which no document states - asked of P2
    ' as REQ-20261009-004 - so no hresult value is emitted yet (never a fabricated one).
    d("errorDecoding") = True
    d("errorTableEntries") = 85
    ' TASK-P1-044: which SHA kernel is ACTUALLY in force, so the forced-fallback evidence
    ' stops being ambiguous. if/else because the earlier one-liner plus a cross-module
    ' Private call did not compile. NOTE: this edit also RESTORES errorTableEntries, which
    ' an earlier block deletion of mine had removed by accident.
    d("shaKernel") = Sha256.ShaKernelName()
    If Sha256.ShaKernelName() = "vb6-fallback" Then
        d("shaFallbackForced") = True
    Else
        d("shaFallbackForced") = False
    End If
    Dim b As Object
    Set b = CreateObject("Scripting.Dictionary")
    b("sha256") = BuildSha256()
    b("builtAt") = BuildBuiltAt()
    ' TASK-P1-036: make the freshness AUDITABLE. sha256/builtAt are read from the
    ' file at call time (same function that feeds instructions.BUILT); path says
    ' WHICH file was read and readAt says WHEN, so a stale value cannot be mistaken
    ' for a live one.
    b("path") = BuildPath()
    b("readAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Set d("build") = b
    Set d("breaking") = BreakingList()
    Set d("migrations") = MigrationsList()
    Set d("changelog") = ChangelogList()
    ' TASK-P1-088: read-only wheel state (pure addition; no new tool, no description change).
    StampWheelNow d
    ToolMcpInfo = JsonStringify(d)
End Function

Private Function BreakingEntry(ByVal since As String, ByVal what As String, ByVal instead As String) As Object
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    e("since") = since
    e("what") = what
    e("instead") = instead
    Set BreakingEntry = e
End Function

Private Function BreakingList() As Object
    Dim c As Object
    Set c = New Collection
    c.Add BreakingEntry("v1.2.0", _
        "get_component_code: coords default disk -> ide (the numbering is now IDE line numbers)", _
        "pass coords=""disk"" for the old text/numbering; the response is a JSON envelope, text in result.text")
    c.Add BreakingEntry("v1.2.0", _
        "replace_code: an out-of-range or continuation edit is REJECTED (range_beyond_end / inside_continuation)", _
        "branch on ok/reason; deletedLines=0 is no longer a success signal, and the IDE content does not change")
    c.Add BreakingEntry("v1.2.0", _
        "write-class tools: validate=off|warn|strict, and edit-side failures carry ok:false + reason", _
        "read reason before retrying; strict refuses the write before the IDE or the disk is touched")
    c.Add BreakingEntry("v1.2.0", _
        "write_file now TRUNCATES; build/make_project/verify report artifactChanged/verdict", _
        "judge a build by artifactChanged/verdict (compiled-no-artifact-change = NOT built), not by ok alone")
    c.Add BreakingEntry("v1.3.0", _
        "version visibility: initialize.instructions banner + per-response mcp:{product,contract} stamp + mcp_info", _
        "purely additive: no existing field, parameter or tool signature changed")
    c.Add BreakingEntry("v1.5.0", _
        "three message-only window tools were added: read_active_dialogs / close_active_dialogs / send_window_message", _
        "purely additive: no existing tool's inputSchema changed. close_active_dialogs never runs automatically and only clicks safe captions; send_window_message only reaches windows of the VB6 process that hosts the add-in")
    c.Add BreakingEntry("v1.6.0", _
        "two tools were added: replace_in_files (cross-component literal/regex replace) and undo_edit (single-step component undo)", _
        "purely additive: no existing tool's inputSchema changed (109 -> 111). replace_in_files defaults to dry_run=true and only ever rewrites the OPEN project; undo_edit restores the component PANE and never touches a file on disk")
    c.Add BreakingEntry("v1.6.0", _
        "send_window_message: HIGH-RISK messages (WM_CLOSE 0x0010 / WM_QUIT 0x0012 / WM_SYSCOMMAND 0x0112) are REFUSED with reason=high_risk_message unless confirm=true", _
        "pass confirm=true only when acting on THAT window IS the intent; nothing was sent when refused. Its parameter is message (an argument named msg is ignored and fails with reason=bad_message) and a raw number (0x0111 / 522 / &H20A) is accepted alongside the named messages")
    c.Add BreakingEntry("v1.8.0", _
        "replace_code with operation=replace or operation=delete now REFUSES the call when expected_sha256 is absent (reason=expected_sha256_required); the hub gained the ide_restart tool (add-in 107 + hub 5 = 112)", _
        "pass expected_sha256 - the refusal carries currentSha256 so it is a single retry; get_component_code reports the same sha. insert/append/set are unchanged. ide_restart only ever stops OUR OWN instance (port >= 8767) and never the external 8766 one")
    c.Add BreakingEntry("v1.6.0", _
        "add_component(type=form): the form is now created by the IDE ITSELF through its own Add Form command, not by a .frm we synthesize", _
        "the receipt carries createdVia: ide_dialog | synthesized (plus ideRouteNote on the fallback) and projectDirWrites is [] on the IDE route; the call waits for the add-in watchdog to press the IDE's own modal confirm button, so it takes about 3.5-4.9 s instead of milliseconds. The synthesized fallback still exists, now with LinkTopic = the component name and the IDE's real default geometry (3015/120/465/4560)")
    c.Add BreakingEntry("v1.6.0", _
        "recorded tool-face limitations: command_bars{list} does NOT enumerate the Add-Ins popup (only execute-by-caption or by id reaches those items), and the AddFromTemplate form route is unavailable on this machine (no plain form template exists in editsec\templates or the VB6 Template\Forms folder) and was deliberately NOT implemented", _
        "reach the settings entry with command_bars{action:""execute"",control:""VB6 MCP Settings...""}; create forms with add_component(type=form) as usual")
    c.Add BreakingEntry("v1.9.0", _
        "three tools were added - immediate_eval / immediate_run / ide_quiet_mode - and ide_state gained release_inflight / release_min_held_ms plus a release{} block (add-in 111 -> 114)", _
        "purely additive for existing callers: no existing tool lost a parameter, and no existing receipt changed on the quiet-OFF / no-release path. ide_state now ALWAYS carries a release{} block with a fixed key set (the same discipline as cancel{}), and read-only receipts gain gate* keys ONLY after a forced release has actually happened")
    ' TASK-P1-089 (contract v1.7.0): the snapshot parameter added by TASK-P1-087 is
    ' ADDITIVE - it is optional and defaults to off, so no existing call changes behaviour
    ' (cross-build A/B over the shared call sequence: 0 differences). This version therefore
    ' adds NO breaking entry ON PURPOSE: fabricating one would misreport the semantics, and
    ' deleting the historical entries above would falsify the record. The v1.7.0 note lives
    ' in MigrationsList() instead.
    Set BreakingList = c
End Function

Private Function MigrationsList() As Object
    Dim c As Object
    Set c = New Collection
    c.Add "line numbers: read coords.default (and lineSpace) in every response before editing; edit tools take IDE line numbers, get_component_code defaults to coords=""ide"""
    c.Add "prefer anchor edits over line numbers: edit_by_anchor / apply_edits with expected_sha256"
    c.Add "call mcp_info once after initialize: product/contract/toolCount + the BREAKING list + how to compare build.sha256"
    c.Add "v1.6.0: replace_in_files defaults to dry_run=true and only rewrites the OPEN project; expected_shas[] makes a mismatching component SKIPPED (never written) and a failure STOPS the run, rolling the applied components back or reporting the exact applied set"
    c.Add "v1.6.0: undo_edit restores a component's PANE text from <project>\_undo\<component>.<ts>.txt (UTF-8). A snapshot is written before every component text write and is gated by VB6MCP_ALLOW_PROJECTDIR_WRITES; to_disk=true is refused (use save_project to flush the project)"
    c.Add "v1.6.0: send_window_message takes message (not msg), accepts raw numbers, and needs confirm=true for WM_CLOSE / WM_QUIT / WM_SYSCOMMAND"
    c.Add "v1.6.0: control_execution.mode's enum now lists the 12 values the implementation really accepts (resume / stop / stepinto / stepover / stepout / runtocursor / fullcompile included as compatibility aliases)"
    c.Add "v1.8.0: replace_code{replace|delete} now REQUIRES expected_sha256 (a line-number edit must pin the text it was measured against; the refusal carries currentSha256 so the fix is one retry) - this is the only caller-visible break. insert/append/set stay optional. validate_code gains an OPTIONAL baseline and every diagnostic gains proc/procOffset/fingerprint/stableKey plus an always-present baseline_diff{added,removed}. ide_state gains cancel_inflight/cancel_target and every read-only receipt carries inflight{}. write-class tools refuse a block that does not pair its own first/last procedure keywords BEFORE writing"
    c.Add "v1.7.0: replace_code / apply_edits / edit_by_anchor accept an OPTIONAL snapshot (boolean, default false; the strings ""on""/""yes""/""1"" are also accepted). snapshot=true guarantees exactly ONE pre-edit pane snapshot so undo_edit can restore that edit; with the argument absent behaviour is UNCHANGED (additive). replace_code's replace/insert/append/delete write no snapshot by default; apply_edits / edit_by_anchor already snapshot by default. undo_edit answers reason=no_snapshot_for_component plus a hint when a NAMED component has no snapshot"
    c.Add "v1.7.0: the frozen tool face was re-frozen once for that additive parameter - tools/schema_baseline.json == docs/tools_schema.json (sha256 1EE3BFA9E917AB6E5D148D409FB6B33ACA91EF2B20C7988E39ADE589536108BE); the tool count is unchanged at 111, and mcp_info.breaking gains NO v1.7.0 entry because nothing in this version breaks a caller"
    c.Add "v1.9.0: immediate_eval / immediate_run run statements in the ACTIVE project's immediate context and carry the VALUE and the PRINTED text out through a private temp file - no Immediate-window pane read, and no Open/Print/Close for the caller. A statement that RAISES still blocks the IDE (On Error is REFUSED on this channel with ret=40036 - measured), so timeout_ms is a GATE-AGE threshold and NOT a hard timeout; immediate_* ARM that release, so the NEXT tool call heals the session by itself, and ide_state{release_inflight:true, release_min_held_ms} forces the gate open on demand"
    c.Add "v1.9.0: ide_state{release_inflight:true} FORCE-RELEASES the single-flight gate of a call that never came back. The receipt reports {released, releasedTool, heldMs, abandonedSeq, statementInterrupted:false} and says the blocked call is STILL RUNNING; the abandoned frame's unwind is swallowed exactly once, so the depth can never be decremented twice. release_min_held_ms (default 60000) keeps a HEALTHY long call from losing the guarantee"
    c.Add "v1.9.0: ide_quiet_mode{on,ttl_ms} is a session-level switch. While ON, control_window{action:""activate""} does NOT call Window.SetFocus (the receipt says quietModeSuppressedFocus:true) and EVERY tools/call receipt carries quiet{...,advisory}. quiet OFF is byte-for-byte the previous behaviour, and foregroundReadable=false means the handle could not be READ (it is 0 here) - never read that as the focus having been left alone"
    Set MigrationsList = c
End Function

Private Function ChangelogList() As Object
    Dim c As Object
    Set c = New Collection
    c.Add "v1.1.0 +5 read-only tools: ide_state / current_position / list_breakpoints / find_gui / ui_text"
    c.Add "v1.2.0 batch A: receipt honesty (coords, sha256 self-report) + write guards"
    c.Add "v1.2.0 batch B+C: 7 new input parameters + 3 read-only tools (workspace_status / verify_scope / todo_status)"
    c.Add "v1.3.0 version visibility: initialize.instructions banner + mcp:{product,contract} stamp + mcp_info; product 0.2.0 -> 0.3.0"
    c.Add "v1.5.0 three message-only tools: read_active_dialogs / close_active_dialogs / send_window_message; command_bars accepts '#<id>' inside its EXISTING control parameter"
    c.Add "v1.6.0 replace_in_files + undo_edit (109 -> 111); send_window_message description completed (the parameter is message, not msg; raw numbers accepted) + high-risk message guard (confirm=true); add_component(type=form) creates the form through the IDE itself (createdVia / ideRouteNote); control_execution.mode enum documented at its 12 accepted values; recorded tool-face limits (command_bars{list} omits the Add-Ins popup; the AddFromTemplate form route is unavailable here)"
    c.Add "v1.8.0 replace_code{replace|delete} requires expected_sha256 (breaking for line-number callers; the refusal returns currentSha256); validate_code baseline + proc/procOffset/fingerprint/stableKey + baseline_diff; ide_state cancel_inflight + inflight{} on every read-only receipt; PRE-WRITE block pairing refusal in replace_code / add_code / apply_edits / edit_by_anchor; multi-line (block) anchors + duplicate-anchor hits + process-scope echo in edit_by_anchor; close_active_dialogs emits one key set on dry_run AND real clicks; add-in 107 + hub 5 = 112 tools (hub ide_restart added); contract v1.7.0 -> v1.8.0"
    c.Add "v1.7.0 optional snapshot parameter (default off) on replace_code / apply_edits / edit_by_anchor + undo_edit structured no_snapshot_for_component; additive, tool count stays 111, frozen tool face re-frozen once (1EE3BFA9...)"
    c.Add "v1.9.0 three tools: ide_quiet_mode (A - foreground safety: SetFocus suppressed, quiet{} + advisory on every receipt), immediate_eval / immediate_run (B - value and printed output carried out through a private temp file, so no pane read and no Open/Print/Close for the caller; an unparseable sink is a structured value_not_parsed); ide_state{release_inflight, release_min_held_ms} + release{} and an ARMED self-healing release (C / defect #10 - the single-flight gate is force-releasable); add-in tool count 111 -> 114; contract v1.8.0 -> v1.9.0"
    Set ChangelogList = c
End Function

' The project may live in a scratch subfolder, so walk up until a _bus\tasks folder
' is found instead of assuming the project directory is the repo root.
Private Function FindBusRoot() As String
    FindBusRoot = ""
    Dim d As String, i As Long, p As Long
    d = ProjectDir()
    If Len(d) = 0 Then d = App.Path
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Function
    For i = 0 To 6
        If Len(d) = 0 Then Exit For
        On Error Resume Next
        If fso.FolderExists(d & "\_bus\tasks") Then
            FindBusRoot = d
            Exit Function
        End If
        On Error GoTo 0
        p = InStrRev(d, "\")
        If p <= 3 Then Exit For
        d = Left$(d, p - 1)
    Next
End Function

Private Function ComponentFileNameOf(ByVal comp As VBIDE.VBComponent) As String
    On Error Resume Next
    ComponentFileNameOf = ""
    ComponentFileNameOf = comp.FileNames(1)
End Function

Private Function AnyObjOf(ByVal args As Object, ByVal key As String) As Object
    On Error Resume Next
    Set AnyObjOf = Nothing
    Set AnyObjOf = args(key)
End Function

Private Function FirstDiagText(ByVal pack As Object) As String
    FirstDiagText = ""
    On Error Resume Next
    Dim arr As Object
    Set arr = pack("diagnostics")
    If arr Is Nothing Then Exit Function
    If arr.Count = 0 Then Exit Function
    Dim d As Object
    Set d = arr(1)
    FirstDiagText = d("ruleId") & " @line " & CStr(d("editLine")) & ": " & d("message")
End Function

' --- minimal JSON field reader for the captured-process JSON (no parser needed) ---
Private Function JsonFieldRaw(ByVal json As String, ByVal field As String) As String
    JsonFieldRaw = ""
    Dim key As String, i As Long, j As Long, ch As String, raw As String
    key = Chr$(34) & field & Chr$(34) & ":"
    i = InStr(1, json, key, vbBinaryCompare)
    If i = 0 Then Exit Function
    i = i + Len(key)
    Do While i <= Len(json) And Mid$(json, i, 1) = " "
        i = i + 1
    Loop
    If i > Len(json) Then Exit Function
    ch = Mid$(json, i, 1)
    If ch = Chr$(34) Then
        j = i + 1
        Do While j <= Len(json)
            ch = Mid$(json, j, 1)
            If ch = "\" Then
                j = j + 2
            ElseIf ch = Chr$(34) Then
                Exit Do
            Else
                j = j + 1
            End If
        Loop
        JsonFieldRaw = Mid$(json, i + 1, j - i - 1)
    Else
        j = i
        Do While j <= Len(json)
            ch = Mid$(json, j, 1)
            If ch = "," Or ch = "}" Then Exit Do
            raw = raw & ch
            j = j + 1
        Loop
        JsonFieldRaw = Trim$(raw)
    End If
End Function

Private Function JsonUnescape(ByVal s As String) As String
    ' TASK-P1-052: was "out = out & ..." per character (quadratic in the decoded text,
    ' which is a git stdout / captured-process JSON field). A pre-sized buffer plus
    ' Mid$ assignment makes it O(n); the decoded text is never longer than the input
    ' (\r \n \t \" \\ \/ and \uXXXX all shrink). Bit-identical: measured on a
    ' 9.5 KB escaped string (same hash) and 475 KB decoded in 32 ms.
    Dim n As Long, i As Long, o As Long, buf As String, ch As String, nx As String
    n = Len(s)
    If n = 0 Then Exit Function
    buf = String$(n, " ")
    i = 1
    Do While i <= n
        ch = Mid$(s, i, 1)
        If ch = "\" And i < n Then
            nx = Mid$(s, i + 1, 1)
            o = o + 1
            If nx = "r" Then
                Mid$(buf, o, 1) = vbCr
            ElseIf nx = "n" Then
                Mid$(buf, o, 1) = vbLf
            ElseIf nx = "t" Then
                Mid$(buf, o, 1) = vbTab
            ElseIf nx = Chr$(34) Then
                Mid$(buf, o, 1) = Chr$(34)
            ElseIf nx = "\" Then
                Mid$(buf, o, 1) = "\"
            ElseIf nx = "/" Then
                Mid$(buf, o, 1) = "/"
            ElseIf nx = "u" Then
                Mid$(buf, o, 1) = ChrW$(CLng("&H" & Mid$(s, i + 2, 4)))
                i = i + 4
            Else
                Mid$(buf, o, 1) = nx
            End If
            i = i + 2
        Else
            o = o + 1
            Mid$(buf, o, 1) = ch
            i = i + 1
        End If
    Loop
    JsonUnescape = Left$(buf, o)
End Function

Private Function ParsePorcelain(ByVal text As String) As Object
    Dim arr As Object
    Set arr = New Collection
    Dim lines() As String, i As Long, ln As String
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 3 Then
            If Left$(ln, 3) <> "## " Then
                Dim d As Object
                Set d = CreateObject("Scripting.Dictionary")
                d("status") = Left$(ln, 2)
                d("path") = Trim$(Mid$(ln, 4))
                arr.Add d
            End If
        End If
    Next
    Set ParsePorcelain = arr
End Function

Private Function BranchOfPorcelain(ByVal text As String) As String
    BranchOfPorcelain = ""
    Dim lines() As String, i As Long, b As String
    Dim p As Long
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        If Left$(lines(i), 3) = "## " Then
            b = Mid$(lines(i), 4)
            p = InStr(b, "...")
            If p > 0 Then b = Left$(b, p - 1)
            BranchOfPorcelain = Trim$(b)
            Exit Function
        End If
    Next
End Function

Private Function CountTag(ByVal text As String, ByVal tag As String) As Long
    CountTag = 0
    Dim p As Long, i As Long, s As String, ch As String
    p = InStr(1, text, tag, vbBinaryCompare)
    If p = 0 Then Exit Function
    i = p + Len(tag)
    Do While i <= Len(text)
        ch = Mid$(text, i, 1)
        If InStr("0123456789", ch) > 0 Then
            s = s & ch
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    If Len(s) > 0 Then CountTag = CLng(s)
End Function

' delta-3 helper: read every .md in a folder and pull the status: line + the H1.
Private Sub CollectMd(ByVal dir As String, ByRef arr As Object)
    On Error Resume Next
    Dim fso As Object, fld As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dir) Then Exit Sub
    Set fld = fso.GetFolder(dir)
    Dim txt As String, ts As Object
    For Each f In fld.Files
        If LCase$(fso.GetExtensionName(f.Name)) = "md" Then
            ' TASK-P1-027 (P4-026 finding): OpenAsTextStream(1, -2) decodes with the system
            ' ANSI code page, so UTF-8 task cards came back as mojibake. Read the BYTES and
            ' decode with the detected encoding instead.
            Dim raw() As Byte
            Encoding.ReadAllBytesTo f.Path, raw
            txt = ""
            Encoding.DecodeBytes raw, Encoding.DetectEncoding(raw), txt
            Dim d As Object
            Set d = CreateObject("Scripting.Dictionary")
            d("file") = f.Name
            d("status") = ExtractLine(txt, "status:")
            d("title") = ExtractH1(txt)
            arr.Add d
        End If
    Next
End Sub

Private Function ExtractLine(ByVal text As String, ByVal prefix As String) As String
    ExtractLine = ""
    Dim lines() As String, i As Long, ln As String
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If LCase$(Left$(ln, Len(prefix))) = LCase$(prefix) Then
            ExtractLine = Trim$(Mid$(ln, Len(prefix) + 1))
            Exit Function
        End If
    Next
End Function

Private Function ExtractH1(ByVal text As String) As String
    ExtractH1 = ""
    Dim lines() As String, i As Long, ln As String
    text = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    lines = Split(text, vbLf)
    For i = 0 To UBound(lines)
        ln = Trim$(lines(i))
        If Left$(ln, 1) = "#" Then
            ExtractH1 = Trim$(Replace(ln, "#", ""))
            Exit Function
        End If
    Next
End Function

' TASK-P1-027: turn "MCPREASON:<reason>|<field>|<field>" into a structured refusal.
Private Function McpReasonJson(ByVal tagged As String, ByVal toolName As String, ByVal args As Object) As String
    Dim rest As String, parts() As String
    rest = Mid$(tagged, 11)
    parts = Split(rest, "|")
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = parts(0)
    d("tool") = toolName
    ' TASK-P1-039 + TASK-P1-046 (contract v1.3.2): vbErrorCode is ONLY for a VB error
    ' GENUINELY captured at the Fail handler. A "MCPREASON:" tag is OUR OWN guard
    ' refusal, so the raise that carries it (GuardSha's Err.Raise vbObjectError+81) is
    ' an implementation detail and must NOT surface as vbErrorCode - otherwise
    ' sha256_mismatch carried -2147221423 and looked like a captured VB error, which is
    ' exactly the tagged-refusal/pre-validation inconsistency P4-038 found. The engine
    ' line is not read here either (this is not an IDE error). IdeBusyBreakJson, which
    ' IS a real IDE error, keeps AddErrorFields.
    ' TASK-P1-071: the field mapping MUST be per-reason. It used to be unconditional - parts(1)
    ' and parts(2) hard-wired to expected/actualSha256 - which was correct only for the single
    ' reason that existed then (sha256_mismatch). A new reason would have handed the caller
    ' MISLEADING KEY NAMES (e.g. a colliding component name reported as "expectedSha256").
    Select Case LCase$(parts(0))
        Case "sha256_mismatch"
            If UBound(parts) >= 1 Then d("expectedSha256") = parts(1)
            If UBound(parts) >= 2 Then d("actualSha256") = parts(2)
        Case "name_collision", "import_path_is_member"
            ' TASK-P1-071: parts(1) = the EXISTING component that caused the refusal,
            ' parts(2) = a suggested action. No sha pair is involved in these refusals.
            If UBound(parts) >= 1 Then d("existingComponent") = parts(1)
            If UBound(parts) >= 2 Then d("suggestedAction") = parts(2)
            Dim srcPath As String
            srcPath = ""
            On Error Resume Next
            srcPath = JsonGetStr(args, "path")
            On Error GoTo 0
            If Len(srcPath) > 0 Then d("sourceFile") = srcPath
            d("sideEffects") = "none: refused before the source snapshot, before AddFile and before SaveAs - the project directory, the component set and the caller's file are unchanged"
        Case "project_dir_unavailable", "project_dir_write_not_allowed"
            ' TASK-P1-077: parts(1) = the suggested action, parts(2) = a one-line explanation.
            ' Neither reason involves a colliding component or a sha pair.
            If UBound(parts) >= 1 Then d("suggestedAction") = parts(1)
            If UBound(parts) >= 2 Then d("detail") = parts(2)
            Dim srcPath2 As String
            srcPath2 = ""
            On Error Resume Next
            srcPath2 = JsonGetStr(args, "path")
            On Error GoTo 0
            If Len(srcPath2) > 0 Then d("sourceFile") = srcPath2
            d("sideEffects") = "none: refused before the source snapshot, before AddFile and before SaveAs - the project directory, the component set, the caller's file and the add-in's own directory are unchanged"
        Case "no_project"
            ' TASK-P1-091: raised by the shared GetProject() helper, so it can reach ANY
            ' project-scoped tool. parts(1) = suggested action, parts(2) = explanation.
            ' No component and no sha pair is involved, and the import wording below would
            ' be wrong here.
            If UBound(parts) >= 1 Then d("suggestedAction") = parts(1)
            If UBound(parts) >= 2 Then d("detail") = parts(2)
            d("sideEffects") = "none: nothing was read, written or modified - there is no open project in the VB6 IDE"
        Case "cancelled_by_request"
            ' TASK-P1-126 (MCP-08): a cooperative abort, raised from a pump loop.
            If UBound(parts) >= 1 Then d("cancelledTool") = parts(1)
            If UBound(parts) >= 2 Then d("cancelledKind") = parts(2)
            d("cancelledByTool") = m_cancelByTool
            d("cancelAgeMs") = GetTickCount() - m_cancelAt
            d("note") = "the in-flight call stopped at its own pump point because a re-entrant READ-ONLY call asked it to (MCP-08). Its own sideEffects/undoSnapshots report any partial work; nothing outside what it had already done was touched."
            d("suggestedAction") = "re-read the component/pane state before retrying, because the aborted call may have completed part of its work. To stop a call blocked inside a HOST modal loop (where none of our pump points run) use the hub tool ide_restart."
        Case "anchor_not_found", "anchor_not_unique", "unknown_op"
            ' TASK-P1-100: an anchor that does not resolve - and an op outside R/I/D - is CALLER
            ' INPUT, so it answers with a branchable refusal instead of the bare +84 raise that
            ' surfaced as ok:false/unexpected_error/vbErrorCode:-2147221420. The tag parts are
            ' NUMBERS ONLY (EngineRefusalTag); the anchor/op/occurrence are read back out of the
            ' caller's OWN args, so arbitrary anchor text never has to survive the "|"-separated
            ' channel. Every read below is guarded: a diagnostic must never turn a refusal into an
            ' unexpected error (TASK-P1-079's state-axis lesson).
            On Error Resume Next
            Dim nfIdx As Long, nfHits As Long, nfPane As Long
            nfIdx = -1
            nfHits = 0
            nfPane = 0
            If UBound(parts) >= 1 Then nfIdx = CLng("0" & parts(1))
            If UBound(parts) >= 2 Then nfHits = CLng("0" & parts(2))
            If UBound(parts) >= 5 Then nfPane = CLng("0" & parts(5))
            If nfIdx >= 0 Then d("failedOpIndex") = nfIdx
            d("hitCount") = nfHits
            If UBound(parts) >= 3 Then d("containsHits") = CLng("0" & parts(3))
            If UBound(parts) >= 4 Then
                If CLng("0" & parts(4)) > 0 Then d("firstContainsLine") = CLng("0" & parts(4))
            End If
            If nfPane > 0 Then d("paneLines") = nfPane
            If UBound(parts) >= 6 Then d("anchorHasCrLf") = (CLng("0" & parts(6)) <> 0)
            d("match") = "normalized"
            d("normalizationApplied") = "LCase + Trim$ (VB6 Trim$ strips SPACES only) + TAB -> space"
            d("noHitAfterNormalization") = (nfHits = 0)
            Dim edList As Object, opItem As Object
            Dim effMch As String
            effMch = ""
            Set edList = AnyObjOf(args, "edits")
            If nfIdx >= 0 Then
                If nfIdx < CollectionCount(edList) Then
                    Set opItem = edList(nfIdx + 1)
                    If Not opItem Is Nothing Then
                        d("anchor") = DictStr(opItem, "anchor")
                        d("op") = UCase$(DictStr(opItem, "op"))
                        If JsonGetLong(opItem, "occurrence", 0) > 0 Then d("occurrence") = JsonGetLong(opItem, "occurrence", 0)
                        effMch = Trim$(DictStr(opItem, "match"))
                    End If
                End If
            End If
            On Error GoTo 0
            If LCase$(parts(0)) = "unknown_op" Then
                Dim allowedOps As Object
                Set allowedOps = New Collection
                allowedOps.Add "R"
                allowedOps.Add "I"
                allowedOps.Add "D"
                Set d("allowedOps") = allowedOps
            End If
            On Error Resume Next
            Select Case LCase$(parts(0))
                Case "anchor_not_found"
                    d("note") = "nothing was written (preview_edit is read-only). The anchor matched 0 of the " & CStr(nfPane) & " pane line(s) under the engine's normalization, so a case / indentation / TAB difference is NOT the cause. See containsHits / firstContainsLine (the anchor is a FRAGMENT of a real line) and anchorHasCrLf (a single anchor is matched against exactly ONE pane line, so an anchor carrying CR/LF can never match)."
                    d("hint") = "call get_component_code{coords:'ide'} and copy the exact single line you want to change"
                Case "anchor_not_unique"
                    d("note") = "nothing was written (preview_edit is read-only). The anchor matched " & CStr(nfHits) & " pane line(s), and preview_edit requires a UNIQUE anchor."
                    d("hint") = "make the anchor unique (widen it), or use apply_edits / edit_by_anchor with an explicit occurrence"
                Case "unknown_op"
                    d("note") = "nothing was written. The op field must be one of R (replace), I (insert after) or D (delete); allowedOps lists them."
                    d("hint") = "set op to R, I or D"
            End Select
            ' TASK-P1-104: with a per-op match="exact" the counts above ARE exact counts, so the
            ' normalized wording would name the wrong cause ("a case difference is NOT the cause"
            ' is false under exact matching - discipline 39). This branch runs ONLY when exact is
            ' in play, so a default-shaped refusal keeps its historical keys and text byte for byte.
            If LCase$(effMch) = "exact" And (LCase$(parts(0)) = "anchor_not_found" Or LCase$(parts(0)) = "anchor_not_unique") Then
                d("match") = "exact"
                d.Remove "normalizationApplied"
                d.Remove "noHitAfterNormalization"
                d("matchApplied") = "binary (vbBinaryCompare) compare of the RAW pane line against the RAW anchor: case and indentation are significant"
                d("noHitWithExactMatch") = (nfHits = 0)
                Select Case LCase$(parts(0))
                    Case "anchor_not_found"
                        d("note") = "nothing was written (preview_edit is read-only). Under EXACT matching the anchor matched 0 of the " & CStr(nfPane) & " pane line(s). Unlike the normalized default, a case / indentation difference IS a possible cause here: drop match=""exact"" to compare normalized. See containsHits / firstContainsLine (a FRAGMENT of a real line, compared RAW here) and anchorHasCrLf (one anchor is matched against exactly ONE pane line)."
                        d("hint") = "copy the exact pane line from get_component_code{coords:'ide'} verbatim (indentation included), or drop match=""exact"" to use the default normalized matching"
                    Case "anchor_not_unique"
                        d("note") = "nothing was written (preview_edit is read-only). Under EXACT matching the anchor matched " & CStr(nfHits) & " pane line(s) and no occurrence was given."
                        d("hint") = "make the anchor unique (widen it) or pass an explicit per-op occurrence"
                End Select
            End If
            On Error GoTo 0
            ' TASK-P1-102: the same per-op audit as the success path, so a caller can see WHICH
            ' per-op key was NOT honoured (TASK-P1-104 moved match/occurrence into the honoured
            ' set, so only a per-op scope can still show up here). Gated on the tool name:
            ' edit_by_anchor answers these reasons with its own receipt, so its shape must not move.
            Dim opkHon As String
            Dim opkMine As Boolean
            opkMine = False
            If LCase$(toolName) = "preview_edit" Then opkMine = True
            If LCase$(toolName) = "apply_edits" Then opkMine = True
            If opkMine Then
                ' TASK-P1-104: preview_edit now resolves through the same pipeline as apply_edits
                ' and honours the same per-op keys, so both tools share ONE honoured set here.
                opkHon = "|op|anchor|text|match|occurrence|"
                On Error Resume Next
                If LCase$(toolName) = "apply_edits" Then
                    OpKeyAuditInto d, AnyObjOf(args, "edits"), opkHon, NoteApplyEditsOpKeys()
                Else
                    OpKeyAuditInto d, AnyObjOf(args, "edits"), opkHon, NotePreviewOpKeys()
                End If
                On Error GoTo 0
            End If
        Case "component_not_found"
            ' TASK-P1-101: FindComponent captured the diagnostics in module state (a component
            ' name can contain anything, so nothing rides the "|" tag). Every read is guarded.
            On Error Resume Next
            d("triedComponent") = m_cnfTried
            If Len(m_cnfTried) > 0 Then
                If Trim$(m_cnfTried) <> m_cnfTried Then d("triedComponentHasSurroundingWhitespace") = True
            End If
            d("componentCount") = m_cnfCount
            If Not m_cnfNames Is Nothing Then Set d("availableComponents") = m_cnfNames
            If m_cnfCount > 100 Then d("availableComponentsTruncated") = True
            If Not m_cnfNear Is Nothing Then
                If m_cnfNear.Count > 0 Then Set d("nearMatches") = m_cnfNear
            End If
            If Len(m_cnfTrimmed) > 0 Then d("trimmedMatch") = m_cnfTrimmed
            If Len(m_cnfCollapsed) > 0 Then d("whitespaceCollapsedMatch") = m_cnfCollapsed
            d("lookupIsCaseInsensitive") = True
            d("caseDifferenceIsNotTheCause") = "the lookup compares component names with LCase$, so a name that differs ONLY in letter case RESOLVES and can never reach this refusal - there is no case fix to make here"
            d("note") = "nothing was read, written or modified: the component you named is not part of the open project, so the tool stopped before touching anything"
            d("hint") = "call get_component_list (or list_component_parts) and pass one of availableComponents verbatim; if trimmedMatch / whitespaceCollapsedMatch is present you passed the right name with different whitespace"
            On Error GoTo 0
        Case "edits_required"            ' TASK-P1-100: BuildOps has exactly two callers (apply_edits, preview_edit).
            d("note") = "nothing was written. The edits argument is missing or empty, so there is no operation to apply or to preview."
            d("hint") = "pass edits as a non-empty array of {op, anchor, text} objects"
        Case Else
            ' an unknown reason keeps only ok/reason/tool - never invent sha keys for it
    End Select
    Dim comp As String
    comp = ""
    On Error Resume Next
    comp = JsonGetStr(args, "name")
    If Len(comp) = 0 Then comp = JsonGetStr(args, "component")
    On Error GoTo 0
    If Len(comp) > 0 And LCase$(parts(0)) <> "no_project" And LCase$(parts(0)) <> "component_not_found" Then
        d("component") = comp
        ' ruling 4: carry coords for any line-numbered field
        On Error Resume Next
        Set d("coords") = CoordsDictOfName(comp, 1, "ide")
        On Error GoTo 0
        d("lineSpace") = "ide"
        d("lineSpaceHint") = "line numbers are IDE code-pane lines; get_component_code's numbered output is DISK lines (ideLine = diskLine - coords.headerOffset)"
    End If
    If d.Exists("sideEffects") Then
        Select Case LCase$(parts(0))
            Case "project_dir_unavailable", "project_dir_write_not_allowed"
                ' TASK-P1-077: this refusal is shared by import_component and by add_component's
                ' form fallback, so the note must not claim it came from the import.
                d("note") = "the refusal happened BEFORE any side effect: no file written into the project directory or into the add-in's own directory, and no component added, rebound or renamed"
            Case "no_project"
                d("note") = "refused before any work: with no project open there is nothing to read or write; open a project and retry"
            Case Else
                ' unchanged wording for the pre-existing reasons (no behaviour drift)
                d("note") = "the import was refused BEFORE any side effect: no snapshot file, no AddFile, no SaveAs, no file in the project directory and no component added or rebound"
        End Select
    ElseIf d.Exists("note") Then
        ' TASK-P1-100: the anchor refusals set their OWN note above. preview_edit writes nothing
        ' at all, so the optimistic-concurrency wording would be WRONG here. No pre-existing
        ' reason sets d("note") before this point, so this branch cannot move their receipts.
    Else
        d("note") = "the component was NOT modified (the optimistic-concurrency guard refused the write)"
    End If
    McpReasonJson = JsonStringify(d)
End Function

' TASK-P1-028: where McpTrace will write (same precedence as the template's EnsureInit).
Private Function SinkPathOf() As String
    SinkPathOf = Environ$("MCP_TRACE_FILE")
    If Len(SinkPathOf) = 0 Then SinkPathOf = Environ$("TEMP") & "\mcp_trace.jsonl"
End Function

' TASK-P1-031: is the sink directory present and WRITABLE? Report it instead of
' letting the trace fail silently at run time (the hardened template creates the
' directory and falls back to %TEMP%, but the caller must be able to see that).
Private Function SinkProbe() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim p As String, dir As String
    p = SinkPathOf()
    dir = ParentDirOf(p)
    d("sinkPath") = p
    d("sinkDir") = dir
    d("sinkDirExists") = False
    d("sinkDirWritable") = False
    If Len(dir) = 0 Then
        d("sinkAdvice") = "the configured MCP_TRACE_FILE has no directory part"
        Set SinkProbe = d
        Exit Function
    End If
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then
        d("sinkAdvice") = "Scripting.FileSystemObject is unavailable, so the sink directory could not be pre-checked"
        Set SinkProbe = d
        Exit Function
    End If
    On Error Resume Next
    d("sinkDirExists") = fso.FolderExists(dir)
    Err.Clear
    Dim tf As Object
    Set tf = fso.CreateTextFile(dir & "\_mcp_probe.tmp", True)
    If Err.Number = 0 Then
        tf.Close
        fso.DeleteFile dir & "\_mcp_probe.tmp", True
        d("sinkDirWritable") = True
    Else
        d("sinkAdvice") = "cannot write into " & dir & " (" & Err.Description & "); the hardened trace module falls back to %TEMP%\mcp_trace.jsonl and logs the failure to MCP_TRACE_ERR_FILE"
    End If
    Err.Clear
    On Error GoTo 0
    Set SinkProbe = d
End Function

Private Function ParentDirOf(ByVal p As String) As String
    Dim i As Long
    i = InStrRev(p, "\")
    If i > 1 Then ParentDirOf = Left$(p, i - 1)
End Function

' ---- TASK-P1-129: automatic modMcpTrace materialization ----

' inject_trace's guard ('If McpTraceOn() Then') has to resolve at COMPILE time, so the
' target project must contain modMcpTrace. Until this card the caller had to import a
' repository file by hand; now the add-in writes the component itself and PROVES it is
' there. The source text and the language detection live in TraceGen (embedded literals).
Private Sub EnsureTraceModule(ByRef info As Object)
    Dim before As Boolean
    before = ModulePresent("")
    info("modulePresentBefore") = before
    info("moduleAutoAttached") = False
    If before Then
        info("moduleAttachAction") = "skipped_already_present"
        Exit Sub
    End If
    Dim langSrc As String
    Dim lang As String
    Dim resLang As String, resWhy As String, uiLang As String, uiWhy As String
    lang = TraceGen.DetectIdeLanguage(langSrc, resLang, resWhy, uiLang, uiWhy)
    info("ideLanguage") = resLang
    info("ideLanguageSource") = resWhy
    info("ideMenuLanguage") = uiLang
    info("ideMenuLanguageSource") = uiWhy
    info("commentLanguage") = lang
    info("commentLanguageSource") = langSrc
    info("commentLanguageRule") = "UI-first (TASK-P1-129 final ruling): the IDE's own UI language wins when it can be read; otherwise the running VB6.EXE version resource; otherwise English."
    Dim code As String
    code = TraceGen.TraceModuleSource(lang)
    info("moduleCodeLines") = (Len(code) - Len(Replace(code, vbCrLf, ""))) \ 2 + 1
    info("moduleCodeSha256") = Sha256.Sha256HexTextUtf8(code)
    info("moduleCodeSha256Note") = "sha256 of the module SOURCE this add-in generated; VB6 may normalize the pane text, so this is the intended source, not the saved .bas"
    Dim t0 As Long
    t0 = GetTickCount()
    Dim addErr As String
    Dim res As String
    addErr = ""
    On Error Resume Next
    res = ToolAddComponent("module", "modMcpTrace", code)
    If Err.Number <> 0 Then
        addErr = "0x" & Hex$(Err.Number) & " " & Err.Description
        Err.Clear
    End If
    On Error GoTo 0
    info("moduleAttachMs") = GetTickCount() - t0
    ' The card requires the attach to be PROVEN, not assumed: re-read the component list
    ' instead of trusting the add call's receipt.
    If ModulePresent("") Then
        info("moduleAutoAttached") = True
        info("moduleAttachAction") = "added"
    Else
        info("moduleAttachAction") = "failed"
        If Len(addErr) > 0 Then
            info("moduleAttachError") = addErr
        ElseIf Len(res) = 0 Then
            info("moduleAttachError") = "add_component returned an EMPTY receipt and modMcpTrace is not in the project"
        Else
            info("moduleAttachError") = "add_component reported no error but modMcpTrace is still not in the project"
        End If
        info("moduleAttachNote") = "the module was NOT attached; the injected guard cannot compile and VB6 will show a MODAL compile-error dialog. This failure is carried in the receipt - nothing is silently downgraded."
    End If
End Sub

' TASK-P1-149: modMcpTrace is NEVER removed automatically. inject_trace attaches it;
' remove_trace removes ONLY that id's injected block. The module is kept because
'   (a) the user may have added his own code to an auto-generated module, and deleting it
'       would destroy that code, and
'   (b) other @mcp-trace blocks may still exist elsewhere and call McpTrace/McpTraceOn,
'       so the project would no longer compile once the module was gone.
' The automatic-removal branch that used to live here (reason "removed_autogen_module") was
' DELETED: there is no code path left that removes the component. Module state is reported
' back through info, and the receipt carries the manual removal instructions.
Private Sub ReportTraceModuleKept(ByRef info As Object)
    info("modulePresentBefore") = False
    info("moduleRemoved") = False
    info("moduleLeftInPlace") = False
    info("moduleAutogenMarker") = False
    info("moduleRemoveReason") = "module_absent"
    Dim comp As VBIDE.VBComponent
    On Error Resume Next
    Set comp = FindComponent("modMcpTrace")
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
    If comp Is Nothing Then
        info("moduleKeptNote") = "no modMcpTrace component exists in this project, so there was nothing to keep or remove"
        Exit Sub
    End If
    info("modulePresentBefore") = True
    info("moduleLeftInPlace") = True
    info("moduleRemoveReason") = "kept_by_design"
    Dim manual As String
    manual = " To remove it yourself: confirm that NO '@mcp-trace' block remains anywhere in the project (search for the marker), then call remove_component with name='modMcpTrace'. remove_trace never does this for you."
    If TraceGen.IsAutogenTraceModule(ComponentPaneText(comp)) Then
        info("moduleAutogenMarker") = True
        info("moduleRemoveNote") = "modMcpTrace is left in place BY DESIGN (TASK-P1-149); the previous automatic removal (reason removed_autogen_module) was withdrawn because (a) you may have added your own code to this auto-generated module and deleting it would have destroyed that code, and (b) other @mcp-trace blocks elsewhere may still call McpTrace/McpTraceOn, so the project would stop compiling once the module was gone." & manual
    Else
        info("moduleRemoveNote") = "modMcpTrace is left in place BY DESIGN (TASK-P1-149). It does not carry the VB6MCP autogen marker, so it is treated as YOUR module - and even an auto-generated one is never removed now, because (a) you may have added your own code to it and (b) other @mcp-trace blocks elsewhere may still call McpTrace/McpTraceOn." & manual
    End If
    Dim boundFile As String
    boundFile = ""
    On Error Resume Next
    boundFile = comp.FileNames(1)
    If Err.Number <> 0 Then
        Err.Clear
        boundFile = ""
    End If
    On Error GoTo 0
    If Len(boundFile) > 0 Then info("moduleFileOnDisk") = boundFile
End Sub

' The code-pane text of one component (empty when it cannot be read).
Private Function ComponentPaneText(ByVal comp As VBIDE.VBComponent) As String
    On Error Resume Next
    ComponentPaneText = ""
    If comp Is Nothing Then Exit Function
    If comp.CodeModule Is Nothing Then Exit Function
    Dim n As Long
    n = comp.CodeModule.CountOfLines
    If n <= 0 Then Exit Function
    ComponentPaneText = comp.CodeModule.Lines(1, n)
End Function

' TASK-P1-031: the injected guard calls McpTrace/McpTraceOn, which only exist when
' the target project contains modMcpTrace. Without it the project does not compile
' and VB6 raises a MODAL compile-error dialog that blocks end/remove_trace - far
' worse than a rejection, so it is pre-checked here.
Private Function ModulePresent(ByVal name As String) As Boolean
    On Error Resume Next
    ModulePresent = False
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If p Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To p.VBComponents.Count
        Dim c As VBIDE.VBComponent
        Set c = p.VBComponents.Item(i)
        If Err.Number = 0 Then
            If LCase$(c.Name) = "modmcptrace" Then
                ModulePresent = True
                Exit Function
            End If
        End If
        Err.Clear
    Next
End Function

' The component file as raw text, for marker searches (empty when there is none).
Private Function DiskTextOfName(ByVal name As String) As String
    On Error Resume Next
    DiskTextOfName = ""
    Dim p As String
    p = DiskPathOf(name, "")
    If Len(p) = 0 Then Exit Function
    If Dir$(p) = "" Then Exit Function
    DiskTextOfName = Build.ReadLogAnsi(p)
End Function

' ---- TASK-P1-032: expression visibility refusal ----

' A MODAL compile-error dialog blocks the whole IDE (end/remove_trace included),
' so inject_trace refuses - structurally - to emit a block that cannot compile.
' The decision is purely lexical (TraceGen symbol scan + Vb6Lint procedure map).
Private Function TraceExprRefusal(ByVal name As String, ByVal text As String, ByVal proc As String, ByVal exprs As String) As String
    Dim vis As String, em As String
    Dim insIdx As Long, afterIde As Long
    vis = TraceGen.VisibleSymbolsAt(text, proc, insIdx, afterIde, em)
    If Len(em) > 0 Then Exit Function
    Dim ok As Object
    Set ok = CsvKeySet(vis)
    AddCsvKeys ok, ComponentNamesCsv()
    AddCsvKeys ok, VbBuiltinNamesCsv()
    Dim items() As String, i As Long
    Dim bad As Object
    Set bad = New Collection
    items = Split(exprs, ",")
    For i = 0 To UBound(items)
        Dim ex As String, tok As String
        ex = Trim$(items(i))
        If Len(ex) > 0 Then
            tok = TraceGen.LeadingIdent(ex)
            If Len(tok) > 0 Then
                If Not ok.Exists(LCase$(tok)) Then
                    Dim o As Object
                    Set o = CreateObject("Scripting.Dictionary")
                    o("expression") = ex
                    o("token") = tok
                    o("why") = "not a parameter of " & proc & ", not declared above the injection point, and not a module-level or project-level symbol"
                    bad.Add o
                End If
            End If
        End If
    Next
    If bad.Count = 0 Then Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "expression_not_visible_at_injection_point"
    d("tool") = "inject_trace"
    d("component") = name
    d("procedure") = proc
    d("insertAfterIdeLine") = afterIde
    Set d("offenders") = bad
    Set d("visibleSymbols") = CsvToCollection(vis, 40)
    d("rule") = "each expression may reference only a parameter of the procedure, a local/Static declared at or above the injection point, or a module-level/project-level symbol; the check is lexical (VB6 is never asked to compile)"
    d("advice") = "move the declaration above the injection point, inject at a later point in the procedure, or drop that expression. Emitting the block unchanged would make the project fail to compile and VB6 would raise a MODAL compile-error dialog that also blocks end and remove_trace."
    d("note") = "the component was NOT modified (the IDE pane, the disk and the pane sha256 are unchanged)"
    TraceExprRefusal = JsonStringify(d)
End Function

Private Function CsvKeySet(ByVal csv As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    AddCsvKeys d, csv
    Set CsvKeySet = d
End Function

Private Sub AddCsvKeys(ByVal d As Object, ByVal csv As String)
    If Len(csv) = 0 Then Exit Sub
    Dim parts() As String, i As Long, k As String
    parts = Split(csv, ",")
    For i = 0 To UBound(parts)
        k = Trim$(parts(i))
        If Len(k) > 0 Then d(LCase$(k)) = True
    Next
End Sub

' a reference to another module/form/class of the project is legal in an expression
Private Function ComponentNamesCsv() As String
    On Error Resume Next
    Dim out As String, i As Long
    Dim p As VBIDE.VBProject
    Set p = GetProject()
    If p Is Nothing Then Exit Function
    For i = 1 To p.VBComponents.Count
        Dim c As VBIDE.VBComponent
        Set c = p.VBComponents.Item(i)
        If Err.Number = 0 Then
            If Len(c.Name) > 0 Then out = out & "," & c.Name
        End If
        Err.Clear
    Next
    If Len(out) > 0 Then ComponentNamesCsv = Mid$(out, 2)
End Function

' VB6 runtime names that may legitimately appear in an expression; without this
' list a valid TypeName(x) / Me.Caption would be refused.
Private Function VbBuiltinNamesCsv() As String
    VbBuiltinNamesCsv = "Me,Len,Left,Left$,Right,Right$,Mid,Mid$,Trim,Trim$,LTrim,LTrim$,RTrim,RTrim$," & _
        "UCase,UCase$,LCase,LCase$,CStr,CInt,CLng,CDbl,CCur,CDate,CBool,CVar,CDec,Chr,Chr$,Asc,Val,Format,Format$," & _
        "Now,Date,Time,Timer,Rnd,Randomize,Int,Fix,Abs,Sgn,Sqr,Exp,Log,Sin,Cos,Tan,Atn,Round,Err," & _
        "UBound,LBound,Array,IsNull,IsEmpty,IsObject,IsNumeric,IsDate,IsArray,IsError,IsMissing," & _
        "TypeName,VarType,IIf,Choose,Switch,RGB,QBColor,Space,Space$,String,String$,Replace,InStr,InStrRev," & _
        "Split,Join,Filter,Hex,Oct,FreeFile,EOF,LOF,Seek,Null,Empty,Nothing,True,False," & _
        "CurDir,Dir,Dir$,FileLen,FileDateTime,GetAttr,SetAttr,MkDir,RmDir,Kill,FileCopy,Name,Shell," & _
        "App,Screen,Printer,Clipboard,Load,Unload,Print,Debug,DoEvents,Command,Environ,Environ$," & _
        "LSet,RSet,TypeOf,AddressOf,New,And,Or,Not,Xor,Mod,Like,Is,ByVal,ByRef,ParamArray,Optional,Object,Variant"
End Function

Private Function CsvToCollection(ByVal csv As String, ByVal maxN As Long) As Object
    Dim c As Object
    Set c = New Collection
    If Len(csv) = 0 Then
        Set CsvToCollection = c
        Exit Function
    End If
    Dim parts() As String, i As Long
    parts = Split(csv, ",")
    For i = 0 To UBound(parts)
        If c.Count >= maxN Then Exit For
        If Len(Trim$(parts(i))) > 0 Then c.Add Trim$(parts(i))
    Next
    Set CsvToCollection = c
End Function
