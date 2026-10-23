Attribute VB_Name = "McpFace"
Option Explicit

' =======================================================================
' A1 of the CodeTools split: the TOOL FACE BUILDER.
'
' Moved byte-for-byte out of CodeTools.cls; the only textual change inside the moved
' bodies is the rename of the two module-level cache variables to explicit ByRef
' parameters (m_tlRootOff -> cacheOff, m_tlRoot -> cacheRoot). CodeTools.cls stays the
' single state host: it declares m_tlRoot/m_tlRootOff and passes them in.
' ToolsCachesReset() (still in the class) clears those class variables, so invalidation
' semantics are unchanged.
' =======================================================================

Public Function BuildToolsListResult(ByRef cacheOff As Long, ByRef cacheRoot As Object) As Object
    If cacheOff = 0 Then
        cacheOff = 1
        If Trim$(Environ$("VB6MCP_TOOLS_CACHE")) = "0" Then cacheOff = -1
    End If
    If cacheOff = 1 Then
        If Not cacheRoot Is Nothing Then
            Set BuildToolsListResult = cacheRoot
            Exit Function
        End If
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    Dim tools As Object
    Set tools = New Collection

    ' ---- contract v1.1.0 batch A: read-only debug/GUI observers ----
    ' (plan B) RETRACTS the earlier promise that a read-only call is served while
    ' another call is in flight. Measured reality: this add-in is single threaded, so a second
    ' call that arrives through the message-pump point of an in-flight call is NOT executed -
    ' it gets the busy refusal receipt (measured: read_active_dialogs got exactly that receipt).
    ' ide_state is therefore NOT a way to work while a call is in flight; its inflight{...} and
    ' cancel{...} blocks only tell you what the refusal refused. The ONLY reliable exits from a
    ' call in flight are (1) WAIT for it to return within its own timeout, or (2) terminate the
    ' VB6 process from outside (ide_restart / Stop-Process).
    Dim pIs As Object
    Set pIs = CreateObject("Scripting.Dictionary")
    Set pIs("cancel_inflight") = Prop("boolean", "BEST EFFORT ONLY (retraction): ask the call currently in flight to STOP. Setting the flag can itself require reaching the same message-pump point, so this is not a way to work while a call is in flight - a second call that arrives there is refused with the busy receipt. When the flag DOES get through, the in-flight call aborts at its next pump point (at least 1500 ms in) with {ok:false, reason:cancelled_by_request}; a call already blocked inside a HOST modal loop reaches none of our pump points and never aborts (measured: 320 s, no return). The only reliable exits are WAIT or terminate the VB6 process from outside (hub ide_restart / Stop-Process).")
    Set pIs("cancel_target") = Prop("string", "Optional: the tool you believe is in flight (informational only - the receipt reports what the add-in actually finds).")
    Set pIs("release_inflight") = Prop("boolean", "FORCE-RELEASE the single-flight gate when the call that holds it never came back. Only acts when the in-flight call is OLDER than release_min_held_ms; the receipt gives {released, releasedTool, heldMs, abandonedSeq, statementInterrupted:false, note}. Safe to send while the gate is held, because this tool is READ-ONLY and therefore never refused by the gate.")
    Set pIs("release_min_held_ms") = Prop("integer", "The age threshold for release_inflight, in ms (default 60000). The gate is NOT touched while the in-flight call is younger than this, so a healthy long call keeps the single-flight guarantee.")
    tools.Add MakeTool("ide_state", _
        "Read the VB6 IDE mode RIGHT NOW (design/run/break) together with the raw VBA6 debug-engine flag, main window title, project name and a running flag. Never cached - a mode switch is visible on the very next call. The receipt ALWAYS carries inflight{active,tool,sinceMs,serviceable,kind} (who is in flight and whether it is a read or a write) and cancel{requested,applied,targetTool,targetKind,sinceMs,reason,note} plus release{released,releasedTool,heldMs,abandonedSeq,statementInterrupted,note}; pass cancel_inflight=true to ASK the in-flight call to stop, or release_inflight=true to FORCE-RELEASE the gate of a call that never came back (only when it is older than release_min_held_ms, default 60000). BOTH ARE BEST EFFORT: while a call holds the message-pump point no second call is executed - the request may not even be delivered - so the reliable exits are WAIT or terminate the VB6 process from outside. " & _
        "It also reports stateSource (ebMode | commandBars | ebMode+commandBars), commandBarsSuspended and notRunning, and clears a detected post-run suspended CommandBars window with ONE activate_component (autoRecover in the receipt) so a following attach call cannot mistake design for run. Waits are resolved EbMode-FIRST: EbMode=2 is break; while the CommandBars model is suspended EbMode=1 is run and EbMode=0 is design (nothing is running), never run. " & _
        "The receipt ALWAYS carries mode (design | run | break | unknown) in EVERY success shape - read that ONE field for the IDE mode: state keeps the raw four-state value (design/run/break/runtime_error) while mode normalises it (runtime_error -> unknown). modeBasis names what decided it (ebMode | commandBars | ebMode+commandBars | runtime_error_dialog | unrecognized_state; present even when nothing is degraded). modeConflict{state,ebMode,expectedFromEbMode} appears ONLY when the raw VBA6 engine flag contradicts the reported state; mode itself always mirrors state.", _
        SchemaWith(pIs))

    tools.Add MakeTool("current_position", _
        "Where the engine is currently stopped, read from the IDE itself: in break mode {available:true, project, module, proc, line, text, ebMode} with line 1-based inside the module and text the raw source line; in design/run mode {available:false, reason:not_in_break} without raising. The raw selection it read is echoed as selectedLine so the claim can be cross-checked.", _
        SchemaWith(Nothing))

    tools.Add MakeTool("list_breakpoints", _
        "List the breakpoints this tool injects into source (marker lines ' @mcp-break <id> begin / Debug.Assert False / end') as {component, line, kind, enabled}. The VBIDE object model has no breakpoint API, so IDE-native breakpoints cannot be enumerated and are reported as native:null instead of being guessed.", _
        SchemaWith(Nothing))

    ' ---- contract v1.1.0 batch C: ui_text ----
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
        "Get detailed information about the open VB6 project(s): name, paths, type, IDE mode (design/run/break), options (description, help file, startup object, start project), components and references. noProject is ALWAYS present: true when the IDE holds no project (the project-scoped fields are then null/empty), false otherwise.", _
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
    Set pOpen("allowHostWrite") = Prop("boolean", "33b: only meaningful with exclusive:true. false (default) = REFUSE to close a DIRTY UNTITLED project (would make VB6 write it into the IDE current directory); true = allow the host write and report it (hostWrites/sideEffects.filesWritten). Named dirty projects are unaffected.")
    tools.Add MakeTool("open_project", _
        "Open a project file (.vbp) in the IDE (VBProjects.AddFromFile). Without exclusive it is added to the current workspace. " & _
        "SAME FILE, SEVERAL SPELLINGS: when the target .vbp is already open - including through an 8.3 short name, forward slashes, different case or a directory alias - the call SHORT-CIRCUITS: it returns alreadyOpen:true with the diskVsMemory comparison and does NOT close or re-add the project. " & _
        "33b: a SUCCESSFUL open always carries ok:true; IDE API errors that did not stop the open are reported in apiWarnings[] (code/message/raw/phase), never in apiError. With exclusive:true, a DIRTY UNTITLED project would be written into the IDE current directory by the close path (measured); the call REFUSES by default with ok:false + reason=dirty_untitled_would_write + dirtyComponents/cwdCandidateFiles/zeroWriteBasis/... and writes nothing - pass allowHostWrite:true to allow that write and get it reported. The refusal depends ONLY on the project-level evidence (VBProject.Saved=false and no file on disk), never on the component list: dirtyComponents is an ARRAY of {name,isDirty,lines?} when the component flags are readable, and null when they are not - dirtyUnknown:true + componentProbe:""unreadable"" + componentSavedReadErrors[] then carry that fact. zeroWriteBasis names the basis of the zero-write claim (normally pre_cwd_scan_equals_post_cwd_scan). ", _
        SchemaWith(pOpen, Req("path")))

    Dim pAddProj As Object
    Set pAddProj = CreateObject("Scripting.Dictionary")
    Dim pAddProj_type As Object
    Set pAddProj_type = Prop("string", "Project type: StandardExe, ActiveXExe, ActiveXDll, ActiveXControl")
    Set pAddProj_type("enum") = EnumList("StandardExe", "ActiveXExe", "ActiveXDll", "ActiveXControl")
    Set pAddProj("type") = pAddProj_type
    Set pAddProj("template_path") = Prop("string", "Optional .vbp template path (VBProjects.AddFromTemplate)")
    tools.Add MakeTool("add_project", _
        "Create a new project in the IDE (VBProjects.Add) or from a template. " & _
        "NOTE: VB6 holds ONE project at a time, so this is an EXCLUSIVE REPLACEMENT: the new project becomes THE project, no project group is created, " & _
        "and UNSAVED changes in the current project are discarded silently - call save_project first.", _
        SchemaWith(pAddProj))

    Dim pRemProj As Object
    Set pRemProj = CreateObject("Scripting.Dictionary")
    Set pRemProj("name") = Prop("string", "Project name (default: active project)")
    Set pRemProj("allowDirty") = Prop("boolean", "false (default) = REFUSE to remove a project with unsaved changes (dirty project or component) because removal discards them; true = discard deliberately. With true, a DIRTY UNTITLED project is still REFUSED unless allowHostWrite:true, because VB6 writes it into the IDE current directory.")
    Set pRemProj("allowHostWrite") = Prop("boolean", "33b: false (default) = REFUSE to close a DIRTY UNTITLED project (VB6 would write it into the IDE current directory); true = allow the host write and report it (hostWrites/sideEffects.filesWritten). Named dirty projects are unaffected.")
    tools.Add MakeTool("remove_project", _
        "Close (remove) a project from the IDE workspace. 33b: a DIRTY UNTITLED project (no file on disk) would be written into the IDE current directory by VB6 while closing (measured: 工程1.vbp/.vbw + component files, same-named files overwritten) and VBProject.Saved is read-only, so the call REFUSES by default with ok:false + reason=dirty_untitled_would_write + dirtyComponents/cwdCandidateFiles/zeroWriteBasis/... and writes NOTHING; pass allowHostWrite:true to allow the write and get it reported in hostWrites/sideEffects.filesWritten. The refusal depends ONLY on the project-level evidence (VBProject.Saved=false and no file on disk): dirtyComponents is an ARRAY of {name,isDirty,lines?} when the component flags are readable and null when they are not - dirtyUnknown:true + componentProbe:""unreadable"" + componentSavedReadErrors[] then carry that fact, and zeroWriteBasis names the basis of the zero-write claim.", _
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
    Set pCode("proc") = Prop("string", "Return only this procedure's range: {startLine,endLine,text,sha256,coords}. startLine/endLine are IDE CODE-PANE lines (the same space every edit tool takes); startLineDisk/endLineDisk give the same range in component FILE lines. The numbered text follows coords, so coords=""disk"" numbers the emitted lines in file space while this range stays in pane space.")
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
    Set pRep("expected_sha256") = Prop("string", "Optimistic-concurrency guard: sha256 of the current pane text. REQUIRED for operation=replace and operation=delete (the call is refused with reason=expected_sha256_required and carries currentSha256, so the fix is a single retry); OPTIONAL for insert/append/set. A line-number edit must pin the text it was measured against. Every call's receipt now carries sha256Before/sha256After/lineDelta plus an explicit ok, so the caller can verify what the pin was compared against.")
    Set pRep("validate") = PropValidate()
    ' opt-in pre-edit pane snapshot. Default OFF, so the hottest editor keeps
    ' exactly its current write cost; pass true when the edit must be rollback-capable.
    Set pRep("snapshot") = Prop("boolean", "true = write the PRE-EDIT pane text into the project's _undo store before this call writes anything (undo_edit can then restore it); default false = no extra file write. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("replace_code", _
        "Edit a component's code in IDE memory (line based). replace: overwrite start_line..end_line. insert: before start_line. append: after last line. delete: remove range. set: replace the whole module with code." & _
        " replace and delete REQUIRE expected_sha256. The chosen remedy is the SHA PIN, not a proc: address space: a whole-pane sha detects ANY drift (a proc-relative address only detects drift across procedure boundaries within the same pane, and it would have to be re-derived from the very text it is meant to pin) - so a drifted pane is refused with sha256_mismatch + the current sha in one retry, and the receipt additionally echoes procBefore/procAfter{proc,startLine,endLine,firstLine,lastLine} so the targeted procedure is visible on both sides." & _
        " (MEASURED; this is NOT an operation property - do not classify by op): whether the pane's final line break survives depends on (a) whether the BASE text already ended with one and (b) WHERE you insert. An insertion in the MIDDLE never touches the tail, so the result INHERITS the base's tail; an insertion at the END rebuilds the tail and normalises it. Measured: same operation (insert, exactly one trailing CRLF submitted) yields result ends_with_crlf=true on a base that already ended with CRLF (middle) and false on a base that did not. Never infer the outcome; the receipt's trailingBlankLine reports what actually happened and newTotalLines counts the way the caller's splitlines() does (a trailing break is not an extra line)." & _
        " the submitted block is checked BEFORE the write - a redundant trailing End <Sub|Function|Property> is refused as extra_procedure_end (with block line numbers, the header/End pairs it did match and a suggestedText), set/append must be self-contained (missing_procedure_end), and any edit that would leave the pane unbalanced is refused as predicted_procedure_imbalance instead of becoming a compile-time E105." & _
        " Component is marked unsaved; use save_project. Pass snapshot=true to write a pre-edit pane snapshot that undo_edit can restore (default false = no extra file write; set already snapshots by default). requestedRange has ONE shape per mode: {mode:""range"", operation, start_line, end_line} here and in add_code / guard refusals, {mode:""anchor"", anchor, occurrence, match, scope?} in edit_by_anchor, {mode:""anchor-batch"", ops, match} in apply_edits - branch on mode instead of guessing which shape arrived.", _
        SchemaWith(pRep, Req("name", "operation")))

    Dim pAddComp As Object
    Set pAddComp = CreateObject("Scripting.Dictionary")
    Dim pAddComp_type As Object
    Set pAddComp_type = Prop("string", "Component type: module, class, form, mdiform, resource, usercontrol, propPage, userdoc, designer")
    Set pAddComp_type("enum") = EnumList("module", "class", "form", "mdiform", "resource", "usercontrol", "propPage", "userdoc", "designer")
    Set pAddComp("type") = pAddComp_type
    Set pAddComp("name") = Prop("string", "Desired component name (default: IDE-generated)")
    Set pAddComp("code") = Prop("string", "Optional initial code (added via AddFromString)")
    Dim pMv As Object
    Set pMv = CreateObject("Scripting.Dictionary")
    Set pMv("from_component") = Prop("string", "Source component name")
    Set pMv("proc") = Prop("string", "Procedure name to move (exact, case-insensitive)")
    Set pMv("to_component") = Prop("string", "Target component name")
    Set pMv("insert_anchor") = Prop("string", "Optional exact line in the target after which the block is inserted; omit to APPEND at the end of the target")
    tools.Add MakeTool("move_procedure", _
        "Move a WHOLE procedure block between components in ONE transaction: take the block from the SOURCE pane, insert it into the TARGET (after insert_anchor, or appended), then delete it from the SOURCE. Any failure rolls BOTH panes back (rolledBack:true, restoreVerified proves the read-back). .vbp is never touched. Receipt: moved{from,to,proc,lineCount,insertAfterLine}, sourceSha256Before/After, targetSha256Before/After, rolledBack, restoreVerified, snapshots[], encoding/roundtrip_ok/bom, vbpTouched:false.", _
        SchemaWith(pMv, Req("from_component", "proc", "to_component")))

    Dim pDp As Object
    Set pDp = CreateObject("Scripting.Dictionary")
    Set pDp("procedures") = Prop("array", "Procedure list: either [{component,proc}, ...] or plain proc names (then component names the single component)")
    Set pDp("component") = Prop("string", "Component used for plain proc-name entries")
    tools.Add MakeTool("delete_procedures", _
        "Delete SEVERAL procedures in ONE transaction: a pre-edit _undo pane snapshot is written for every touched component (undo_edit can restore it); if any deletion fails every touched pane is restored. Receipt: deleted[]{component,proc,fromLine,toLine,lineCount}, snapshotId, rolledBack, restoreVerified, shas[].", _
        SchemaWith(pDp, Req("procedures")))

    Dim pBa As Object
    Set pBa = CreateObject("Scripting.Dictionary")
    Set pBa("path") = Prop("string", "Absolute or project-relative file path")
    Set pBa("old_bytes_b64") = Prop("string", "Base64 of the exact ORIGINAL bytes to replace (GBK/ANSI bytes as stored on disk)")
    Set pBa("new_bytes_b64") = Prop("string", "Base64 of the replacement bytes (may be empty to delete, or contain an INSERTION)")
    Set pBa("all_matches") = Prop("boolean", "default false = the anchor must occur exactly once; true = replace every occurrence")
    tools.Add MakeTool("edit_by_byte_anchor", _
        "Replace bytes in a file using GBK/ANSI RAW-BYTE anchors (base64 in, base64 out) - the only byte-exact edit path for mixed Chinese/ASCII sources. Refuses an anchor that would split a double-byte GBK character (lead 0x81..0xFE / trail 0x40..0xFE) with reason=anchor_splits_gbk_char; reports encoding_mismatch when the bytes are valid text but the file is not in that encoding. A round-trip self-check (decode->encode) runs BEFORE writing and a failure refuses the write (roundtrip_ok:false). Receipt: encoding, bom{...}, gbkBoundarySafe, anchorByteOffset, matches, sha256Before/After, bytesBefore/After, roundtrip_ok.", _
        SchemaWith(pBa, Req("path", "old_bytes_b64", "new_bytes_b64")))

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
    ' item 2, discipline 84: the description MUST stop naming
    ' AddFromFile - the file branch no longer uses it, and the observable behaviour changed
    ' (the designer head is stripped, and the target component is no longer renamed).
    Set pAddCode("file") = Prop("string", "Source file whose BODY is appended: the designer head (VERSION/BEGIN..END or Begin..End, plus the leading Attribute lines) is stripped first, so the code pane never receives a header and the target component is never renamed. Supply code or file - at least one of the two is required")
    Set pAddCode("validate") = PropValidate()
    tools.Add MakeTool("add_code", _
        "Append code to a component's code module from a string or a file. Requires code or file. (MEASURED BEHAVIOUR): whether the module's final line break survives is NOT a property of this tool or of the operation - it depends on the base text's tail and, for an insertion, on the position (middle inherits the tail, end rebuilds it). Do not infer it: the receipt's trailingBlankLine reports the ACTUAL result, and newTotalLines counts as the caller's splitlines does. inDeclarationArea is the DECLARATION-AREA verdict, DEFINED as ""the block landed BEFORE the first procedure header line"" (where Option Explicit / module-level Dim/Const/Type live) - NOT ""outside an unclosed procedure body"": on a module that already has procedures an END append is therefore false, and an insert above the first procedure header is true. insertedAt{ide,disk} is the MEASURED landing in both line spaces.", _
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
    Set pFind("start_line") = Prop("integer", "Search from line (1-based, default 1). ADVANCE SEMANTICS: to read the next hit of the same target, pass the previous receipt's nextStartLine (= the matched line + 1). A start_line past the last line is the END of the walk and answers found:false + reason=no_more_hits; the search NEVER wraps back to line 1, so a page cannot silently return a hit above start_line.")
    Set pFind("start_col") = Prop("integer", "Search from column (1-based, default 1)")
    Set pFind("end_line") = Prop("integer", "Search up to line (default: last line)")
    Set pFind("end_col") = Prop("integer", "Search up to column (default: last column)")
    Set pFind("whole_word") = Prop("boolean", "Whole word match only")
    Set pFind("match_case") = Prop("boolean", "Case sensitive")
    Set pFind("pattern_search") = Prop("boolean", "Use pattern matching")
    Set pFind("include_procedure") = Prop("boolean", "ADDITIVE, default false. When true every hit also carries the IDE's own procedure mapping for the hit line: proc (the enclosing procedure; an empty string when the line belongs to no procedure), procBodyLine and procStartLine (IDE code-pane lines) plus procSource=ide.ProcOfLine. With the flag absent the receipt is key-for-key unchanged.")
    Set pFind("include_leading_comment") = Prop("boolean", "ADDITIVE, default false. When true every hit also carries leadingComment: the contiguous block of single-quote comment lines DIRECTLY ABOVE the hit's procedure header (no blank line between the block and the header) plus leadingCommentLines. Implies the procedure mapping. With the flag absent the receipt is key-for-key unchanged.")
    tools.Add MakeTool("find_in_component", _
        "Find text inside one component using CodeModule.Find. Returns the match position (start/end line/column) plus hits[] (0 or 1 element), count and nextStartLine. PAGINATION: pass the previous nextStartLine as start_line; when there is no further hit the receipt answers found:false with reason=no_more_hits (start_line>1, i.e. a continuation page) or reason=not_found (a first page) and hits[]=[] / count=0. A start_line past the last line is not an error: it is the end of the walk. The result never falls back to a hit before start_line and never invents a hit in another procedure - a hit reported outside the requested window is discarded and the receipt says so.", _
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
    ' the \`view\` property was REMOVED from this schema. VB6EXT v5.3 declares
    ' CodePane.CodePaneView GET-ONLY (the type library put-map of _CodePane holds only
    ' TopLine) and has no \`View\` property at all, and a REAL-IDE scan of all 34 command
    ' bars found NO view toggle - so the enum full_module/procedure was a promise no
    ' add-in can keep. A caller who still passes \`view\` now gets ignoredArgs.
    Set pSetCp("view") = Prop("string", "NOT a supported argument: VB6's object model exposes CodePaneView as READ-ONLY and no command bar carries a view toggle, so the value is IGNORED and reported back in ignoredArgs (declared so a caller who passes it gets that answer instead of unknown_argument)")
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
    Set pWrite("create_dirs") = Prop("boolean", "Create missing parent directories (default: true). The directories actually created are listed in the receipt as createdDirs[] and counted by createdDirsCount; if a parent cannot be created the call answers ok:false + reason=parent_dir_missing + missingPath (reasonLegacy=parent_not_creatable) and writes nothing.")
    Set pWrite("backup") = Prop("boolean", "Keep a .bak copy of the previous file (default: true)")
    Set pWrite("validate") = PropValidate()
    Set pWrite("confirm_truncate") = Prop("boolean", "default false. Guard against ACCIDENTAL data loss: when the target ALREADY EXISTS and is NON-EMPTY and content is empty/blank, the write is REFUSED with reason=empty_content_would_truncate. Pass true to intentionally empty a NON-VB file. VB source files (.bas/.cls/.frm/.ctl/.pag/.dsr/.vbp) are ALWAYS refused, even with confirm_truncate:true - an empty VB source cannot compile; delete the file explicitly if that is really intended.")
    tools.Add MakeTool("write_file", _
        "Write a text file to disk with the given encoding. For files of the open project use reload_component afterwards (or replace_code for in-memory edits). VB project/source files (.vbp/.frm/.cls/.bas/.ctl) are written ANSI without BOM; never hand-write them as UTF-8-with-BOM (VB6 reads a mangled first keyword and the project stops loading).", _
        SchemaWith(pWrite, Req("path", "content")))

    Dim pSearch As Object
    Set pSearch = CreateObject("Scripting.Dictionary")
    Set pSearch("query") = Prop("string", "Text or regular expression to search for")
    Set pSearch("path") = Prop("string", "Directory to search; relative to the project directory when not absolute (default: project directory)")
    Set pSearch("pattern") = Prop("string", "Semicolon-separated file patterns (default: VB source files)")
    Set pSearch("is_regex") = Prop("boolean", "Treat query as a regular expression (default: false)")
    Set pSearch("case_sensitive") = Prop("boolean", "Case-sensitive matching (default: false)")
    Set pSearch("max_results") = Prop("integer", "Maximum number of matches (default: 200)")
    Set pSearch("include_procedure") = Prop("boolean", "ADDITIVE, default false. When true every result also carries a LEXICAL procedure mapping for the hit line: proc (the enclosing procedure; empty when the line is outside every procedure), procBodyLine (header line in the FILE), procLineSpace=file and procSource=text_scan - search_in_files has no IDE CodeModule for an arbitrary file, so this is a text scan, not the IDE's ProcOfLine. With the flag absent the receipt is key-for-key unchanged.")
    Set pSearch("include_leading_comment") = Prop("boolean", "ADDITIVE, default false. When true every result also carries leadingComment: the contiguous block of single-quote comment lines directly above the enclosing procedure's header (no blank line between) plus leadingCommentLines. Implies the procedure mapping. With the flag absent the receipt is key-for-key unchanged.")
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
    Set pMake("project") = Prop("string", "Optional ABSOLUTE path of ANY .vbp to compile (route=cli only). When set, that project is compiled even if a DIFFERENT project is open in the IDE; a CLI build registers a CLSID snapshot around the /make and restores/ reports it. route=ide refuses this parameter (the IDE route compiles whatever the IDE holds).")
    Set pMake("timeout_ms") = Prop("integer", "route=cli: bounds ONLY the VB6.EXE /make subprocess (default 180000)")
    Set pMake("allow_host_modal") = Prop("boolean", "Default false. When VB6's own Make Project box is already up, confirming it starts a long host-modal compile that this single-threaded add-in cannot interrupt - every later write tool would then answer reason=busy until the IDE is restarted. false = leave the box alone and report modal_blocked=true with its hwnd/title/controls instead of clicking it.")
    Set pMake("laa") = Prop("boolean", "default TRUE. After the compile, set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header. The add-in rewrites the PE header ITSELF, in place, and re-reads it to verify; an external EditBin.exe is NOT required. The receipt always reports laaBefore/laaAfter/laaSet (+ laa{} with the detail); an artifact that already carries the flag is left untouched (laaSet=false, laaAlready=true). The decision PRECEDENCE is: this per-call laa argument > the [VB6MCP] section of the project .vbp (LAA=1, optional LAANote=...) > the optional sidecar <vbpDir>\.vb6mcp.json {""laa"": true|false} > the environment variable VB6MCP_LAA > the global default (ON). The receipt reports laaSource as param | vb-project | project-file | env | global-default (plus laaProjectFile). An unreadable/invalid sidecar is NOT silently defaulted: the step is skipped and laaError/laaSource=project-file-error is returned." & _
            " WHERE the switch is stored: the named section [VB6MCP] of the .vbp is the PRIMARY source (level 2). Survival rule: a named section present WHEN THE PROJECT IS LOADED is re-emitted by VB6 on every save and survives; a section APPENDED externally WHILE the project is open is overwritten by that save - so write the .vbp with the IDE CLOSED, or use the sidecar file, which can be written at any time. It is the ROOT-area keys that the in-memory model overwrites (CondComp is that class), NOT a named section.")
    Set pMake("wait_for_build") = Prop("integer", "Milliseconds this call waits for the artifact to land before reporting buildState=pending (default 0 = the historical 15000 ms; a smaller value is raised to 15000; cap 120000). The wait is what makes buildState trustworthy: a pending build is NEVER reported as completed. After clearing a host modal box you must re-read output{exists,size,mtime,sha256} - clearing the box is not completion.")
    Set pMake("diagnose") = Prop("boolean", "default false. Run an ISOLATED VB6.EXE /make over a COPY of the .vbp plus its component files in a temp directory and return the real compiler log plus firstError{file,line,text}. This is the only route that can diagnose a project the IDE holds (an isolated /make on the HELD project exits 0 and writes no log). It runs AUTOMATICALLY when the in-IDE Make returns compiled=true but artifactChanged=false, because that verdict is otherwise indistinguishable from a compile failure.")
    tools.Add MakeTool("make_project", _
        "Compile the project. NEVER deletes an existing artifact: the previous output is moved aside first and restored if the build fails or if the host deletes it (output{exists,size,mtime,sha256} + artifact_restored report every case, so a failed build can no longer leave you with no artifact). compiled:true alone is NOT evidence that anything was built - read compiled_means, artifactChanged and verdict. The default ide route writes NO compiler log, so log/logPath/errorCount/firstError are always present with logAvailable=false plus the reason; route=cli returns the build tool's receipt, which carries the real log, errorCount and firstError{file,line,text}." & _
        " After a successful compile the artifact's PE header gets IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) written by this add-in itself (laa, default true; no external EditBin.exe is needed) and the receipt reports laaBefore/laaAfter/laaSet; diagnose=true compiles a COPY of the project in a temp directory with an isolated VB6.EXE /make and returns a real log plus firstError (it also runs automatically when the in-IDE Make says compiled=true but artifactChanged=false, a verdict that is otherwise indistinguishable from a compile failure)." & _
        " the receipt carries ONE queryable build state - buildState (blocked_by_host_modal | pending | completed | completed_without_artifact_change | failed | no_artifact), derived from the same facts as verdict. pending means the compile started but the artifact had not landed inside this call's wait window: re-read output{exists,size,mtime,sha256} (and confirm MZ) before treating it as done - compiled/ok are never completion evidence, and clearing a host modal box is not either. wait_for_build=<ms> extends this call's own wait (default 0 = 15000 ms, cap 120000); the receipt reports waitRequestedMs/waitUsedMs.", _
        SchemaWith(pMake))

    Dim pExec As Object
    Set pExec = CreateObject("Scripting.Dictionary")
    Dim pExec_mode As Object
    Set pExec_mode = Prop("string", "run (start, or continue when already broken) | full_compile | break (pause) | continue | resume (= continue) | step_into | step_over | step_out | run_to_cursor | end | stop (= end) | restart (end then run). Aliases accepted for compatibility: stepinto, stepover, stepout, runtocursor, fullcompile.")
    Set pExec_mode("enum") = EnumList("run", "full_compile", "break", "continue", "resume", "stop", "step_into", "step_over", "step_out", "run_to_cursor", "end", "restart")
    Set pExec("mode") = pExec_mode
    Dim pExec_allow As Object
    Set pExec_allow = Prop("boolean", "default false (INFORMED CONSENT). Starting from DIRTY IDE memory (VBProject.Saved=False or a component IsDirty=True) is REFUSED with reason=" & Chr$(34) & "dirty_memory_would_be_saved" & Chr$(34) & " unless this is true; the refusal carries the dirty component list and howToProceed. dirtyUnknown (flags unreadable) never refuses. This does NOT disable the IDE option - it only makes the caller acknowledge that a Start may write the IDE memory to disk.")
    Set pExec("allowDiskMutation") = pExec_allow
    tools.Add MakeTool("control_execution", _
        "Control program execution in the VB6 IDE: start/continue, start with full compile, break, single-step (into/over/out), run to cursor, end or restart. Locale independent: menu items are matched by caption, ignoring the localized bar name, across English and Chinese IDEs - including the Continue item that replaces Start while the program runs or is broken. " & _
        "An end/stop that finds NOTHING running (EbMode=0 and design) returns ok:true + state_not_changed:true + alreadyNotRunning:true instead of ok:false: in the post-run suspended CommandBars window the End item is disabled precisely because the program has already ended." & _
        " R-1021-02: every receipt carries diskWrites[{file,beforeBytes,afterBytes,beforeSha,afterSha,reason,attribution}] + touchedFiles[] + diskSnapshot{hashMode,budgetBytes,checkedFiles,snapshotMs,taken} + saveOnRun + ideAutoSaveHint (the IDE-wide option Tools > Options > Environment > 'When a program starts' may save the in-memory edits on any Start; this tool never changes it). A Start from DIRTY memory is refused unless allowDiskMutation:true." & _
        " END/STOP while the program RUNS is supported (measured path: the End commandbar returns true in RUN state and really stops the program; the old reason=commandbars_unavailable_while_running refusal for end/stop was a FALSE rejection). Such a receipt carries endCommandDelivered / endPending / stateAfterVerifiedBy=ide_state / stateAfterReadBack / programEnded / deliveryVsStateNote, and programEnded=true appears ONLY after a design read-back - an issued-but-not-yet-observable End answers ok:false + endPending=true + howToContinue.", _
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
    Set pCb("bar") = Prop("string", "Optional bar filter: a bar name OR a menu path segment (matched as a caption prefix), e.g. """ & ChrW$(&H8C03) & ChrW$(&H8BD5) & """ or ""Debug""")
    Set pCb("control") = Prop("string", "Required for action=execute. A menu command caption (prefix search allowed), or an id as ""#746"" / ""id:746"" / a bare number, optionally ""#746 <caption>"" to disambiguate.")
    tools.Add MakeTool("command_bars", _
        "List the IDE menus (CommandBars) or execute a menu command. action=list returns every bar and control with its language-independent identity (id, tag, type, shortcutText, enabled, visible, path) AND TWO EXECUTABILITY FIELDS per control: executable (whether action=execute would actually run a command for that caption right now) and enabledReason (why not, or why a disabled control is still executable). Never infer executability from enabled alone: captions repeat across bars and toolbar buttons, and execute matches by caption and takes the first ENABLED match, so a listed enabled=false control can still be executable through an enabled twin and a listed enabled=true one can be unreachable when every twin of its caption is disabled." & _
        " List the IDE menus (CommandBars) or execute a menu command. action=list returns every bar and control with its language-independent identity (id, tag, type, shortcutText, enabled, visible, path). action=execute runs one command, addressed either by caption or by id as ""#746"" / ""id:746"" / a bare number; because VB6 ids are NOT unique, an id-only lookup executes only when EXACTLY ONE ENABLED control matches and otherwise returns a structured id_ambiguous failure with idCandidates - disambiguate with bar= or by appending the caption (""#746 <caption>""). The CommandBars model is suspended while a program runs, so use this in design or break state.", _
        SchemaWith(pCb, Req("action")))

    ' ---- contract v1.5.0: message-only dialog primitives ----
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
        & "A second, independent guard: every path - the dry_run preview included - carries host_modal_risk (none|possible|known) + host_modal_risk_reason + host_modal_risk_class + riskDialogs[], and a box whose button would start a host-modal operation is NOT clicked unless allow_host_modal=true; a Cancel/Close caption is still allowed on such a box, because cancelling cannot start the operation. recovery_hint explains what to do if a write tool then answers reason=busy. Every path returns the same keys: clickedCount, dialogsBefore, dialogsAfter, clicked[].stillExists, skipped[], remaining[].", _
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
        "Quit the VB6 IDE (VBE.Quit) with an HONEST receipt. A process cannot observe its own exit (once it exits, no receipt can be written), so exited is ALWAYS false here and the exit is confirmed by an observer OUTSIDE this process. If anything BLOCKS the quit - unsaved project/component changes, or a modal box that owns the IDE (ownerDisabled=true) - NOTHING is requested: ok:false + requested:false + blockedBy + evidence{dialogs[],unsavedProjects[],unsavedComponents[]} + howToClear, and the IDE is left running and unchanged. Otherwise the quit IS requested: ok:true + requested:true + pendingExit:true + waitedMs (the measured pre-check) + howToConfirm. quitting:true is kept only as a legacy alias of requested, never as a claim of exit. The quit runs about 150 ms AFTER this response is sent, so confirm the exit outside: poll ide_state / ide_list / the instance registry until THIS pid is gone." & _
        " 34b (ADDITIVE, receipt-only): a REQUESTED quit also carries exitWrites{declaration,observed,condition,evidence,basis}. It is a DECLARATION, not a measurement: this tool cannot see its own exit phase and therefore cannot confirm, block or undo a write there. declaration is always advisory_only and observed is always not_observable. condition is decided ONLY from the project count visible now: 0 -> no_project_observed_no_write (measured: nothing was written, 63 samples); 2 or more -> project_group_context_then_clean_quit (measured: a .vbg appeared, 50-53 B, about 250 ms AFTER the receipt);" & _
        " exactly 1 -> undetermined, because the two measured shapes DISAGREE (a single-project clean quit wrote nothing, while a group-context quit that held 1 project wrote 50 B), so the count alone cannot tell them apart. evidence is a FIXED 4-entry list (W37.T6 written, W38b.B1b written, W38b.B1a not written, W38.S5 not written) and basis names those windows plus 8-hex sha prefixes." & _
        " sideEffects.filesWritten covers ONLY writes performed DURING the call: it never covers the exit phase and must not be read as a guarantee about it. A BLOCKED quit (requested:false) carries NO exitWrites key at all.", _
        SchemaWith(Nothing))

    Dim pVal As Object
    Set pVal = CreateObject("Scripting.Dictionary")
    Set pVal("name") = Prop("string", "Component name (kind inference + reporting)")
    Set pVal("code") = Prop("string", "Code text to validate (NOT written to the IDE)")
    Set pVal("kind") = Prop("string", "bas|cls|frm (optional; inferred from name)")
    Set pVal("mode") = Prop("string", "off|warn|strict (default warn)")
    Set pVal("baseline") = Prop("string", "a previous run to diff against - either the stableKey lines or the WHOLE previous receipt text (both are scanned for stableKey values). The receipt then reports baseline_diff{added[],removed[]} matched on stableKey = ruleId|message|proc, so inserting lines and shifting every line number does NOT make a pre-existing diagnostic look new.")
    tools.Add MakeTool("validate_code", _
        "Validate VB6 code (lexical + structural + agent-habit rules) before writing; returns diagnostics with stable keys, line numbers and sha256. Every diagnostic carries proc + procOffset + fingerprint + stableKey, and baseline_diff{added[],removed[],unchangedCount} is always present (enabled=false unless baseline is given)." & _
        " linter != compiler. These rules are LEXICAL + STRUCTURAL only - they never resolve types or symbols, so a CLEAN receipt is NOT a compile pass and this tool can never report a call to a method that does not exist (that needs type information; a cheap approximation would be guessing) - an explicit NON-GOAL. For a real compile verdict use build (isolated VB6.EXE /make, real compiler log + firstError{file,line,text}) or make_project.", _
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
    Set pCc("baseline") = Prop("string", "Optional previous check_component/validate_code receipt text (or just its stableKey lines). DECLARED and honoured: the receipt then carries baseline_diff{enabled:true,added[],removed[],unchangedCount}. Fail-closed: text that carries neither a 'stableKey' nor a pipe-delimited key line is refused with reason=baseline_unusable, because it cannot be compared.")
    tools.Add MakeTool("check_component", _
        "Lint the current in-IDE text of a component (same diagnostics as validate_code, evaluated on IDE memory). baseline=<a previous receipt or its stableKey lines> makes the receipt compare against it: baseline_diff.enabled=true with added[]/removed[] matched on stableKey. A baseline that cannot be compared is refused with reason=baseline_unusable - never reported as 'no differences'. Unknown argument names are refused with reason=unknown_argument + acceptedArguments." & _
        " linter != compiler - lexical + structural only, no type resolution, so it can never report a call to a method that does not exist (NON-GOAL) and a clean receipt is not a compile pass. Use build for a real compiler log and firstError{file,line,text}.", _
        SchemaWith(pCc, Req("name")))

    Dim pB As Object
    Set pB = CreateObject("Scripting.Dictionary")
    Set pB("project") = Prop("string", "Project .vbp path (optional; default = active project)")
    Set pB("defines") = Prop("string", "Comma-separated /d defines, e.g. TESTING=1")
    Set pB("out_path") = Prop("string", "Build log path (optional)")
    Set pB("timeout_ms") = Prop("integer", "Timeout in ms (default 180000)")
    Set pB("laa") = Prop("boolean", "default TRUE. After the compile, set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header. The add-in rewrites the PE header ITSELF, in place, and re-reads it to verify; an external EditBin.exe is NOT required. The receipt always reports laaBefore/laaAfter/laaSet (+ laa{} with the detail); an artifact that already carries the flag is left untouched (laaSet=false, laaAlready=true). The decision PRECEDENCE is: this per-call laa argument > the [VB6MCP] section of the project .vbp > the optional sidecar <vbpDir>\.vb6mcp.json > the environment variable VB6MCP_LAA > the global default (ON); the receipt reports laaSource (param | project-file | global-default | env) and laaProjectFile, and an invalid sidecar is reported as project-file-error instead of being silently defaulted.")
    tools.Add MakeTool("build", _
        "Compile the project with an isolated VB6.EXE /make <vbp> /out <log> subprocess (no modal) - the reliable command-line route. Returns log (the compiler log TEXT), logPath proven readable via logAvailable/logBytes, errorCount, firstError{file,line,lineSpace,text}, output{exists,size,mtime,sha256}, and artifact_locked + artifact_holder_pid when the artifact was not rewritten because a running process holds it (instead of a bare permission-denied). Note: only the FIRST error is reported per build. laa (default true) makes the add-in set IMAGE_FILE_LARGE_ADDRESS_AWARE (0x0020) in the artifact's PE header itself (no external EditBin.exe) and the receipt reports laaBefore/laaAfter/laaSet.", _
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
    ' same opt-in knob as replace_code (this tool already snapshots through the
    ' write funnel by default; snapshot=true keeps that guarantee explicit and writes no
    ' second file).
    Set pEba("snapshot") = Prop("boolean", "true = guarantee a PRE-EDIT pane snapshot exists in the project's _undo store (this tool already writes one by default through the write funnel); default false = unchanged. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("edit_by_anchor", _
        "Replace one uniquely-matched line by content anchor (no line numbers). The anchor may SPAN MULTIPLE LINES - a pasted block matches when all its lines match consecutively (normalized by default) and the WHOLE matched block is replaced; when the anchor is missing or ambiguous the refusal carries hits[] with the candidate {line, proc, text} of every match (plus coords/lineSpace), so pick one with occurrence=<n> or narrow it with scope instead of guessing." & _
        " a new block that carries a redundant trailing End <Sub|Function|Property> is refused BEFORE the write (reason=extra_procedure_end, with the offending block line numbers, the pairs it did match, and a suggestedText), and an edit that would leave the pane unbalanced is refused as predicted_procedure_imbalance - instead of becoming a compile-time E105." & _
        " the receipt echoes the affected procedure as procBefore/procAfter{proc,startLine,endLine,firstLine,lastLine}. snapshot=true is accepted (a pre-edit pane snapshot for undo_edit is written by default here).", _
        SchemaWith(pEba, Req("name", "anchor", "new_text")))

    Dim pAe As Object
    Set pAe = CreateObject("Scripting.Dictionary")
    Set pAe("name") = Prop("string", "DEFAULT component name. Optional when every edits[] item carries its own component; if neither is present the call is refused with reason=component_required + opIndex.")
    Set pAe("edits") = Prop("array", "Array of {op:R|I|D, anchor, text, component?, match?, occurrence?}. MULTI-COMPONENT: each item may name its own component; the tool-level name is the default. Ops apply in edits[] order and each anchor is resolved against ITS OWN component UPDATED text (earlier ops may move later anchors); components are independent. Nothing is written unless every op resolves and every touched component writes (see rolledBack/restoreVerified).")
    Set pAe("expected_sha256") = Prop("string", "Optional guard for the DEFAULT/single component; the whole batch is all-or-nothing")
    Set pAe("expected_shas") = Prop("array", "Optional pre-write guard for multi-component batches: [{component, sha256}]. ANY mismatch REFUSES the whole batch with guards[{component,expected,actual,ok}] and writes nothing; the legacy expected_sha256 still guards the default/single component.")
    Set pAe("max_ops") = Prop("integer", "Batch size limit (default 200, capped at 1000); a batch with more ops, or more than 40 distinct components, is refused with ok:false reason=batch_too_large + limits{ops,components,max_ops} and writes nothing.")
    Set pAe("match") = Prop("string", "normalized (default) or exact")
    Set pAe("validate") = PropValidate()
    ' same opt-in knob as replace_code (this tool already snapshots through the
    ' write funnel by default; snapshot=true keeps that guarantee explicit and writes no
    ' second file).
    Set pAe("snapshot") = Prop("boolean", "true = guarantee a PRE-EDIT pane snapshot exists in the project's _undo store (this tool already writes one by default through the write funnel); default false = unchanged. The JSON boolean true is expected; the strings on/yes/1 are also accepted")
    tools.Add MakeTool("apply_edits", _
        "Apply a batch of anchor edits atomically (any failure leaves the component unchanged). Every op takes the same multi-line block anchors as edit_by_anchor (a block match replaces the whole matched block), and a failing op answers with failedOpIndex + hits[] {line, proc, text} for candidate selection (occurrence/scope, or per-op match). Each op's new text is checked for a redundant trailing End <Sub|Function|Property> before anything is written, and a batch that would leave the pane unbalanced is refused as predicted_procedure_imbalance. The receipt echoes procBefore/procAfter for the affected procedure. snapshot=true is accepted (a pre-edit pane snapshot for undo_edit is written by default here). requestedRange here is {mode:""anchor-batch"", ops, match}; the same field name carries {mode:""range"", operation, start_line, end_line} on replace_code/add_code and {mode:""anchor"", ...} on edit_by_anchor - branch on mode." & _
        " MULTI-COMPONENT BATCH (v1.16.0): one call can carry ops across up to 40 components (edits[].component, default = name) - ONE pre-edit _undo snapshot per touched component, ONE rollback unit, and NO new tool. Order semantics: ops apply in edits[] order and every anchor is resolved on its own component UPDATED text (earlier ops may move later anchors); components are independent. " & _
        " All-or-nothing: on ANY failure every component already written is restored to its pre-batch pane text and the restore is PROVEN by read-back (rolledBack + restoreVerified + restoredComponents[]), then ok:false + failedOpIndex (GLOBAL index) + hits[]. A successful multi-component receipt adds appliedOps / totalOps / components[{component,ops,shaBefore,shaAfter,snapshotId,verified}] / batchElapsedMs / guards / limits; a single-component call keeps the legacy receipt key-for-key. Pre-write guards: expected_shas[{component,sha256}] (and the legacy expected_sha256) refuse the batch with guards[] and write nothing. Limits: max_ops (default 200, capped 1000) and 40 components -> reason=batch_too_large.", _
        SchemaWith(pAe, Req("edits")))

    Dim pPe As Object
    Set pPe = CreateObject("Scripting.Dictionary")
    Set pPe("name") = Prop("string", "DEFAULT component name (optional when every edits[] item carries its own component)")
    Set pPe("edits") = Prop("array", "Array of {op:R|I|D, anchor, text, component?, match?, occurrence?}. MULTI-COMPONENT: each item may name its own component; the tool-level name is the default. Hunks are reported GROUPED PER COMPONENT. Nothing is written.")
    Set pPe("max_ops") = Prop("integer", "Preview batch size limit (default 200, capped at 1000); more than 40 components is refused with reason=batch_too_large.")
    Set pPe("match") = Prop("string", "Batch matching mode: normalized (default) or exact - the SAME top-level switch apply_edits honours. A per-op match overrides it for that op.")
    tools.Add MakeTool("preview_edit", _
        "Return a unified-style diff for a batch of anchor edits without writing anything. It walks the SAME resolution as apply_edits (per-op match/occurrence, MULTI-LINE block anchors that replace the whole matched block) and applies the same structural refusals, so previewing and then writing the same batch cannot disagree. B+C batch (contract invariant 8): the algorithm is a LINE-LEVEL LCS with hunk grouping, so an insertion in the middle renders as ONE hunk of +N (not a delete block plus an add block); contextLines (default 3) of context are emitted around each hunk. hunks[] carries {kind,start,count,oldStart,oldCount,newStart,newCount,contextLines}; 'algorithm' reports lcs | fallback-too-large | none, and the fallback is stated instead of silently degrading. Still a preview: nothing is written." & _
        " MULTI-COMPONENT PREVIEW (v1.16.0): the same edits[] shape as apply_edits (edits[].component, default = name) across up to 40 components in ONE call; the answer groups hunks per component (components[{component,ops,hunks,hunkCount,algorithm,...}] + totalHunks) and it still writes NOTHING. Order semantics match the write: ops in edits[] order, each anchor resolved on its component UPDATED text. Limits: max_ops (default 200, capped 1000) / 40 components -> reason=batch_too_large.", _
        SchemaWith(pPe, Req("edits")))

    Dim pWs As Object
    Set pWs = CreateObject("Scripting.Dictionary")
    Set pWs("path") = Prop("string", "Target file path")
    Set pWs("text") = Prop("string", "Text to write (ASCII string literals only for VB files)")
    Set pWs("encoding") = Prop("string", "auto|ansi|utf-8|utf-16le (default auto: VB files -> ansi)")
    Set pWs("validate") = PropValidate()
    Set pWs("confirm_truncate") = Prop("boolean", "default false. An EXISTING NON-EMPTY target plus empty/blank text is REFUSED with reason=empty_content_would_truncate; VB source files (.bas/.cls/.frm/.ctl/.pag/.dsr/.vbp) are ALWAYS refused even with confirm_truncate:true (an empty VB source cannot compile).")
    tools.Add MakeTool("write_source", _
        "Write a VB source/text file with explicit encoding and byte-level sha256 (does not touch the IDE). The parent directory is created when missing and reported as createdDirs[]/createdDirsCount; a parent that cannot be created is a structured refusal (ok:false + reason=parent_dir_missing + missingPath + reasonLegacy=parent_not_creatable) and nothing is written. Every success is READ BACK from the target path and must match the intended sha256 byte for byte - the receipt carries exists/size/sha256/verifiedOnDisk and a write that cannot be read back is never reported as written. VB project/source files (.vbp/.frm/.cls/.bas/.ctl) are written ANSI without BOM; never hand-write them as UTF-8-with-BOM (VB6 reads a mangled first keyword and the project stops loading).", _
        SchemaWith(pWs, Req("path", "text")))

    Dim pVf As Object
    Set pVf = CreateObject("Scripting.Dictionary")
    Set pVf("name") = Prop("string", "Component name")
    Set pVf("disk_path") = Prop("string", "Optional disk file (default = component file)")
    Set pVf("baseline") = Prop("string", "Optional lint baseline keys (newline separated)")
    Set pVf("requireCompiler") = Prop("boolean", "false (default) keeps this receipt byte-for-byte. true DEMANDS COMPILER evidence: an isolated VB6.EXE /make over a COPY of the project on disk must produce a real log. When the compiler cannot be resolved, cannot be started, writes no log or is killed at the budget, the call is REFUSED (ok:false, structured reason, compilerProbe{resolved,path,why}, howToProceed) and a lint verdict is NEVER substituted. On success the receipt adds authority=compiler + compiler{path,logPath,errorCount,artifactChanged,compiledFrom}. This path never saves the project and never touches the real artifact.")
    tools.Add MakeTool("verify", _
        "Aggregate verdict for one component: IDE==DISK + lint diagnostics + optional baseline filter. " & _
        "requireCompiler:true demands COMPILER authority instead of the cached/lint answer: an isolated VB6.EXE /make over a COPY of the project on disk must produce a real log; when no compiler is available the call is REFUSED (ok:false + reason=compiler_not_configured + compilerProbe{resolved,path,why} + howToProceed) and a lint verdict is NEVER substituted. A successful compile adds authority=compiler + compiler{path,logPath,errorCount,artifactChanged}. Default false keeps the receipt exactly as before.", _
        SchemaWith(pVf, Req("name")))
    ' W-H4: ONE-call debug session (face 110 -> 111; refreeze deferred to the window).
    Dim pDs As Object
    Set pDs = CreateObject("Scripting.Dictionary")
    Set pDs("runArgs") = Prop("string", "Command line arguments passed to the started program (optional)")
    Set pDs("env") = Prop("object", "{NAME: value} environment variables for the started program (optional)")
    Set pDs("breakpoints") = Prop("array", "Array of {component, line (PANE line space), condition?, hitCount?, enabled?}")
    Set pDs("actions") = Prop("array", "Array of {op: step_over|step_into|step_out|eval|continue|stop, expr? (eval only)}")
    Set pDs("timeout_ms") = Prop("integer", "Hard budget for the WHOLE session in ms (default 60000)")
    Set pDs("attach") = Prop("boolean", "true = drive an ALREADY running/broken IDE (no Start click); false (default) = arm + start and return startedPending")
    Set pDs("watch") = Prop("array", "Optional expressions evaluated at every break/step and merged into transcript[].watch as [{expr, value?, typeName?, error?}]. ONLY identifiers enumerated by an earlier get_locals in this IDE session are evaluated; an unknown name is recorded as {expr, error:""unknown_identifier_refused""} and NEVER evaluated (an unknown identifier can hang the in-flight call - measured). Each entry is individually wrapped in a gate backstop.")
    tools.Add MakeTool("debug_session", _
        "Drive ONE whole VB6 debug session inside a SINGLE call: arm the given breakpoints, START the program WITHOUT saving the project, " & _
        "wait for a breakpoint hit, then execute the ordered actions (step_over/step_into/step_out/eval/continue/stop) and return transcript[] " & _
        "{seq,kind,position{module,proc,paneLine,diskLine},proc,callstack,locals,output,err} plus stopReason/wedged. The VB6 IDE is single-threaded " & _
        "(STA), so a debug session cannot be driven by concurrent calls - it is compressed into this one call. This tool NEVER saves the project " & _
        "(preRunSave=false: the run path deliberately bypasses control_execution, whose run does a pre-run flush); call save_project first if the " & _
        "project must be on disk. The session REQUESTS control_execution(end) and then waits a BOUNDED 3 s inside the call; because End is a QUEUED command in a single-threaded IDE, that wait cannot observe it - a state that still reads break/run is reported as endQueued=true + howToFinish, NOT as wedged. " & _
        "wedged=true means the call could not recover the IDE in-call: an EbMode=2 break context held while the CommandBars model is gone, and one activate_component attempt did not clear it. EVERY return path (normal/timeout/error) also reports diskVbpShaAfter/savedDuringCall/reorderedDuringCall. Refused with " & _
        "reason=already_in_break when the IDE is already broken (end it first). Every breakpoint is echoed with armed/paneLine/diskLine/hitCount, " & _
        "and the .vbp is read from DISK before and after (diskVbpShaBefore/diskVbpShaAfter) - equality is the 'nothing was written' proof. " & _
        "The IDE may still be finishing the queued end/start commands after this call returns: confirm with ide_state and call control_execution{end} if needed (howToFinish is provided when that is the case).", _
        SchemaWith(pDs, Req()))

    ' ---- W-H1: run_until_break (face 111 -> 112; refreeze deferred to the window) ----
    Dim pRub As Object
    Set pRub = CreateObject("Scripting.Dictionary")
    Set pRub("timeout_ms") = Prop("integer", "Budget for the OBSERVATION phase in ms (default 30000, capped at 600000). The gate backstop is armed for this budget + 8 s so a read-only call can force-release a call that never returns.")
    Set pRub("attach") = Prop("boolean", "false (default) = arm + START and return startedPending immediately (Start is a QUEUED command in the STA IDE); true = OBSERVE the already running/broken IDE until break / process exit / unhandled error / timeout")
    Set pRub("breakpoints") = Prop("array", "Array of {component, line (PANE line space), hitCount?} armed before starting (attach=false only). hitCount N means: stop on the Nth hit, auto-continue the earlier hits.")
    Set pRub("watch") = Prop("array", "Optional expressions evaluated at every break and merged into transcript[].watch; ONLY identifiers enumerated by an earlier get_locals are evaluated, an unknown name is recorded as {expr, error:""unknown_identifier_refused""} and never evaluated.")
    tools.Add MakeTool("run_until_break", _
        "Run the project and stop at one of FOUR outcomes: break / process exit / unhandled error / timeout. " & _
        "TWO-PHASE because the VB6 IDE is single-threaded (STA) and Start is a QUEUED command: call ONE with attach=false - it arms breakpoints, clicks Start " & _
        "and returns startedPending=true + howToContinue (the program starts only AFTER the call returns); call TWO with attach=true - it OBSERVES and returns " & _
        "the UNIFIED debug receipt {ok,state,position,proc,callstack,locals,output,err,timeoutMs,wedged,stopReason,hitCountTotal,transcript[]} plus " & _
        "diskVbpShaBefore/diskVbpShaAfter and savedDuringCall/reorderedDuringCall (the debug path never saves). hitCount>1 is the engine-independent equivalent " & _
        "of a conditional breakpoint: every hit still stops the engine, the driver auto-continues the first N-1 hits and stops on the Nth (VB6 has NO conditional " & _
        "breakpoints; expression conditions are NOT supported and are not accepted). The single-flight gate is backstopped with ArmGateRelease(timeout_ms + 8 s): " & _
        "if this call never returns, a later read-only call can force-release it.", _
        SchemaWith(pRub, Req()))

    ' ---- W-H1: breakpoints (face 112 -> 113; refreeze deferred to the window) ----
    Dim pBp As Object
    Set pBp = CreateObject("Scripting.Dictionary")
    Set pBp("action") = Prop("string", "list | set | clear | enable | disable | clear_all | probe (default list). probe is the BEHAVIOURAL verifier and it is TWO-PHASE: phase=start only STARTS the program (a start command is QUEUED in the STA IDE and takes effect AFTER the call returns), phase=observe (same probeId) waits for the stop, reads position/callstack, ends the program and PROVES the state by read-back. It never saves the project and never writes the source.")
    Set pBp("component") = Prop("string", "Component name (set/clear/enable/disable/probe)")
    Set pBp("anchor") = Prop("string", "For mode=injected: the source line to break before (must match uniquely inside a procedure body)")
    Set pBp("line") = Prop("integer", "PANE line number (1-based) for mode=native (and the position to clear); for probe it is the TARGET line whose breakpoint effectiveness is observed")
    Set pBp("mode") = Prop("string", "injected (default: a source-marker Debug.Assert block, survives the session, rewritten by the IDE only through this tool) | native (Debug-menu toggle on the caret line: NO source write, but the VBIDE has no breakpoint API so its state cannot be READ back)")
    Set pBp("id") = Prop("string", "Stable handle (default: generated)")
    Set pBp("clearStale") = Prop("boolean", "false (default). mode=native: when a STALE ledger record exists for this component+line (paneSha changed), clear it and proceed with the F9 post instead of refusing with stale_native_record.")
    Set pBp("timeout_ms") = Prop("integer", "breakpoints{action:probe}: bounded observation budget in ms (default 15000, capped 120000); for mode=native it is the bound of the async F9 channel observation. Ignored by other actions.")
    Set pBp("phase") = Prop("string", "probe only: start (default) = run the preconditions, START the program and return probeId + probeResult=started_pending (a start command is QUEUED in the STA IDE, so the run begins AFTER this call returns); observe = with that probeId, wait up to timeout_ms for the stop, read position/callstack, END the program and PROVE stateAfter=design by read-back (unknown/expired probeId => structured refusal reason=probe_not_found); auto = explicit alias of start (NOT both phases in one call - a single call cannot observe the run it starts).")
    Set pBp("probeId") = Prop("string", "probe phase=observe: the probeId returned by THIS instance phase=start. Single-use; unknown or older than 600000 ms => ok:false + reason=probe_not_found + howToProceed.")
    Set pBp("hitCount") = Prop("integer", "Stop on the Nth hit (default 1). Enforced by run_until_break / debug_session, not by the engine; a conditional EXPRESSION is not supported (VB6 has no conditional breakpoints)")
    tools.Add MakeTool("breakpoints", _
        "Manage REAL IDE breakpoints by component + source anchor / PANE line: set (injected source marker OR native Debug-menu toggle), clear, enable, disable, list, clear_all. " & _
        "ENUMERATION SOURCE is declared on every receipt: the list merges (a) source-marker breakpoints this add-in injected (readable again through list_breakpoints) and (b) native breakpoints set BY THIS ADD-IN's session (tracked in memory only); IDE-native breakpoints set by a HUMAN inside the IDE cannot be enumerated or verified because the VBIDE object model has NO breakpoint API - that limit is reported, never guessed. " & _
        "Every entry carries BOTH coordinate spaces: paneLine (the line an edit tool takes) and diskLine (paneLine + the component's SharedHeaderOffset), plus the coordinate-source note. " & _
        "hitCount is the engine-independent equivalent of a conditional breakpoint (stop on the Nth hit, auto-continue the earlier hits); expression conditions are NOT supported and not accepted. " & _
        "enable/disable for mode=native toggles the IDE breakpoint (its on/off state cannot be read back, so the receipt says nativeStateVerified:false and relies on the session record). " & _
        "action=probe is the only BEHAVIOURAL verifier and it is TWO-PHASE. phase=start (default) checks design state / no modal box / breakable line, STARTS the program (never saves it), then returns probePhase=start, probeResult=started_pending, probeId, startedPending=true and howToContinue. phase=observe (same probeId) waits up to timeout_ms for the stop and reports probeResult=stopped_at_target | stopped_elsewhere (with position{component,proc,paneLine} + callstack), or timeout | did_not_stop | did_not_stop_unobserved (with programStillRunning measured), and it ENDS the program and PROVES stateAfter=design by read-back (stateAfterVerifiedBy=ide_state) - it never claims an end it did not observe. A start command is QUEUED in the STA IDE, so a single call cannot observe the run it starts: phase=auto is an explicit alias of start and a positive result can only come from start + observe." & _
        " The observe HIT path can be a TWO-STAGE wrap-up: when the End command is delivered but its teardown is not observable inside the call, the receipt answers endPending=true + howToContinue and the SAME phase=observe with the SAME probeId re-reads the state; programEnded=true appears ONLY in that confirm receipt. phase=start always carries startQueued=true (the Start command is queued - a measured fact, not a promise that the program is already running). " & _
        "THREE RULES stated explicitly: (1) paneLine is the coordinate space EVERY edit tool takes (diskLine = paneLine + SharedHeaderOffset is DISK space); " & _
        "(2) breakpoints can only be set in DESIGN or BREAK state - while the program RUNS the CommandBars model is suspended, so action=set cannot reach the Debug menu; " & _
        "(3) action=probe does NOT read the anchor parameter - it needs component + line (anchor belongs to mode=injected action=set only).", _
        SchemaWith(pBp, Req()))

    ' ---- W-H2: get_locals / eval_at_frame (face 113 -> 115; refreeze deferred) ----
    Dim pGl As Object
    Set pGl = CreateObject("Scripting.Dictionary")
    Set pGl("frame_index") = Prop("integer", "0 = innermost frame (default). frame_set binds the VBA6 evaluation frame before every read.")
    Set pGl("allowUnsafeInBreak") = Prop("boolean", "default false. In BREAK state get_locals is REFUSED with reason=break_eval_requires_opt_in: the value reads evaluate identifiers in the IDE immediate context and a raising evaluation can wedge the single-threaded pipe (measured). true is INFORMED CONSENT - it does NOT guarantee the statement will not hang, and a hang is recovered with the hub-side ide_kill_and_relaunch (NOT taskkill). A refusal never arms the gate; every failure receipt carries gateReleased + howToRecover.")
    tools.Add MakeTool("get_locals", _
        "Read the CURRENT procedure parameters and declared locals in BREAK mode, with their engine values. " & _
        "SOURCE DECLARATION: variable NAMES come from the SOURCE of the current procedure (parameter list plus Dim/Static declarations in its body); the VALUES are read by binding the VBA6 evaluation frame (vba6_probe frame_set path) and evaluating each name in the IDE immediate context - localsSource is echoed on every receipt. " & _
        "Nothing is invented: when no value can be read the receipt returns localsUnavailable:true + localsUnavailableReason + the raw frame diagnostics. " & _
        "Each value carries typeName + value (or error) and, for arrays/Dictionary/objects, a SHALLOW inspection with explicit depthLimit=2 / widthLimit=20, truncatedDepth, truncatedWidth and cycleDetected (ObjPtr identity). VB6 has no reflection, so a custom class reports its TypeName/ObjPtr and fieldsUnavailable:true rather than guessed fields. " & _
        "The per-value shape is the SAME one transcript[].watch uses (one helper).", _
        SchemaWith(pGl, Req()), _
        " RISK + OPT-IN: reading locals in BREAK mode can wedge the pipe (single-threaded STA, non-preemptible statement). It is REFUSED unless allowUnsafeInBreak:true; the refusal carries howToWorkAround + howToRecover (prefer the hub-side reset ide_kill_and_relaunch over taskkill). The intended safe path for a one-shot read of the break context is debug_session (one call, gate backstopped).")

    Dim pEaf As Object
    Set pEaf = CreateObject("Scripting.Dictionary")
    Set pEaf("expr") = Prop("string", "VB expression to evaluate (e.g. i*2)")
    Set pEaf("frame_index") = Prop("integer", "0 = the IDE's current frame (default). OUTER frames are refused: reaching them needs the engine frame-set thunk, which is FORBIDDEN because it terminated the VB6 process (measured).")
    Set pEaf("allowUnsafeEval") = Prop("boolean", "false (default). In BREAK (and RUN) the immediate context is NOT the break frame and a raising expression can hang the in-flight call (measured), so evaluation there is REFUSED unless this is true. When true the call is wrapped in ArmGateRelease(timeout_ms) and a receipt proves the gate was released; recovery for a hang = ide_state{release_inflight:true, release_min_held_ms:...} and, if needed, restart the IDE.")
    Set pEaf("timeoutMs") = Prop("integer", "CALLER BUDGET HINT in ms (default 5000). The add-in CANNOT preempt a VB statement that is already executing: if this budget passes without a return, the CALLER must time out and use howToRecover (ide_state{release_inflight:true, release_min_held_ms:5000}, restart the IDE if needed). The internal ArmGateRelease is armed for this value + slack regardless.")
    Set pEaf("timeout_ms") = Prop("integer", "Alias of timeoutMs (kept for compatibility)")
    tools.Add MakeTool("eval_at_frame", _
        "Evaluate an expression in the IDE immediate context with an EXPLICIT SAFETY CONTRACT. " & _
        "frameBinding is ALWAYS ideImmediateContext and it is NOT the break frame: measured on the live bed, a*2 read 0 while a had been assigned 7 in the broken frame. The VBA6 frame-set thunk is FORBIDDEN here (measured: it terminated the VB6 process / pipe_stalled) - OUTER frame_index is refused, never attempted. " & _
        "In DESIGN the expression is evaluated normally. In BREAK (and RUN) evaluation is REFUSED with reason=eval_in_break_unsafe + howToUnblock, because a raising expression can hang the in-flight call (measured); pass allowUnsafeEval=true to execute anyway - the call is then wrapped in ArmGateRelease(timeout_ms, default 5000) and the receipt proves gateReleased. " & _
        "howToRecover is always present: ide_state{release_inflight:true, release_min_held_ms:<ms>} frees a hung single-flight gate; restart the IDE if the process itself is wedged. " & _
        "For object/array/Dictionary values a shallow inspection is returned ONLY in a safe context (design or allowUnsafeEval): depthLimit=2, widthLimit=20, truncatedDepth, truncatedWidth, cycleDetected. A failure NEVER returns a fabricated value: stopped_at + raw diagnostic text instead. " & _
        "CALLER BUDGET CONTRACT: timeoutMs (default 5000) is a HINT - the add-in cannot preempt an executing VB statement, so if it does not return within that budget the caller must stop waiting and recover with ide_state{release_inflight:true, release_min_held_ms:5000} (restart the IDE if needed). ANY state can hang an evaluation (design included: Array(1,2,3) hung in a measured run), so do NOT wait on a 120 s transport budget.", _
        SchemaWith(pEaf, Req("expr")))

    ' ---- W-H3: run_project / last_error (face 115 -> 117; refreeze deferred) ----
    Dim pRp As Object
    Set pRp = CreateObject("Scripting.Dictionary")
    Set pRp("commandLineArgs") = Prop("string", "Arguments the started program should see as Command$. LIMITATION (documented): a Standard EXE started from the VB6 IDE does NOT receive Command$; only the ActiveX Component start mode has Debugging-tab arguments. When this is non-empty the tool REFUSES to start (commandLineArgsSupported:false) instead of pretending - build the EXE and run it with args, or use an ActiveX project.")
    Set pRp("env") = Prop("object", "{NAME: value} environment variables applied to the IDE process BEFORE Start (same process as the IDE-run program, so Environ() sees them). Previous values are captured and restored when the program ends; envApplied[] lists every one.")
    Set pRp("startMode") = Prop("string", "Standalone (default) | ActiveX Component. NOTE: the start mode is a PROJECT property and cannot be switched at run time - the receipt echoes the project's current startMode and says startModeApplied:false when it differs.")
    Set pRp("observe") = Prop("boolean", "false (default) = PHASE 1: apply env, click Start, return startedPending + howToContinue; true = PHASE 2: observe until break / exit / error dialog / timeout, then restore env")
    Set pRp("timeout_ms") = Prop("integer", "Observation budget in ms (default 15000, capped 120000)")
    Set pRp("observe_max_wait_ms") = Prop("integer", "OPTIONAL observe budget in ms (default: timeout_ms, clamped to 1000..120000). This is a REAL early return: when it expires the call answers state=""timeout"" + observedState + elapsedMs and RETURNS, so the single-flight gate is not kept. Only meaningful with observe:true.")
    Set pRp("done_marker") = Prop("object", "OPTIONAL {path, contains}: finish observing as soon as the file AT PATH ON THE MACHINE THAT RUNS THE IDE (UNC allowed) contains the case-sensitive sequence. A pure-ASCII value is matched as bytes; a non-ASCII value is tried as UTF-8 and as system ANSI bytes (marker.matchedBy = utf8 | ansi). Only the first 65536 bytes are read and NOTHING is written. On a hit: state=""done-by-marker"" + marker{path,contains,matched,bytesRead,matchedBy,caseSensitive}.")
    Dim pRp_allow As Object
    Set pRp_allow = Prop("boolean", "default false (INFORMED CONSENT). Starting from DIRTY IDE memory is REFUSED with reason=" & Chr$(34) & "dirty_memory_would_be_saved" & Chr$(34) & " unless this is true; the refusal carries the dirty component list and howToProceed. dirtyUnknown never refuses. This does NOT disable the IDE option - it only makes the caller acknowledge that a Start may write the IDE memory to disk.")
    Set pRp("allowDiskMutation") = pRp_allow
    tools.Add MakeTool("run_project", _
        "Start the OPEN project from the VB6 IDE with control over ENVIRONMENT and (where supported) command-line arguments, in TWO PHASES (Start is a QUEUED command in the STA IDE): phase 1 applies env (SetEnvironmentVariable, BEFORE Start - the IDE and the IDE-run program are the same process, so Environ() reads it) and returns startedPending + howToContinue; phase 2 observes and returns {startedPid, runningState, firstBreakpointHit, err}. " & _
        "env is restored when the program ends (envApplied[]/envRestored). VB6 CANNOT attach to an -Embedding process that is already running: the supported path is always to start a debuggable instance from the IDE. " & _
        "A Standard EXE started from the IDE does NOT receive Command$ (the Debugging-tab arguments belong to the ActiveX Component start mode), so a non-empty commandLineArgs is REFUSED with commandLineArgsSupported:false instead of silently ignored. The project start mode is a project property and is echoed, never switched at run time." & _
        " R-1021-02: the phase-1 receipt carries diskWrites/touchedFiles/diskSnapshot/saveOnRun/ideAutoSaveHint (the IDE option may save the in-memory edits on Start) and refuses a dirty Start unless allowDiskMutation:true; phase 2 (observe) starts nothing, so its diskSnapshot.taken is false.", _
        SchemaWith(pRp, Req()), _
        " observe is BOUNDED (observe_max_wait_ms, default timeout_ms) and can also end at a done_marker file on the IDE machine; the receipt carries state (done-by-marker | exited | break | error | timeout - a CLOSED set of FIVE, mutually exclusive) and an ADDITIVE diag{...} block (blockingProc/clockImpl/clockStart/clockAtLoopEntry/clockAtExit/tickDelta/iterations/errNumber/waitSpans/clockSelfTest, plus clockSuspect when this build clock readings do not advance) + observedState + elapsedMs + marker{...} + pipeReleasedByReturn:true. The budget is enforced INSIDE the wait loop: a timeout receipt is ok:true + timeout:true + retrySafe:true + suggestedAction=observe_again with the MEASURED elapsedMs (NOT a failure, NOT the goal being reached) and, when a marker was requested but never " & _
            "matched, marker.unmetReason=""budget reached before the marker appeared"". done-by-marker is emitted ONLY when marker.matched:true; a program that ends normally is state=exited, and when a marker was requested but never matched the receipt says marker.unmet:true (""the run finished " & _
            "without producing the evidence you asked for"" - do NOT read it as success). observe in BREAK state is REFUSED with reason=observe_in_break_not_supported + howToProceed (continue/end the program first), and a #32770 dialog is CLASSIFIED before it can be called an error: an IDE dialog (Find/Replace/Options/...) answers reason=observe_blocked_by_ide_dialog + dialogPrecheck{dialogs[],isProgramError:false} and is NEVER reported as a program error. " & _
        " STA QUEUEING (measured live): while the program RUNS the IDE message pump does NOT dispatch a new tool call, so this call is served only after the program yields (DoEvents), hits a breakpoint, or exits. observe in RUN state is therefore NOT real-time observation: it means wait until the program yields/ends, and state=timeout is structurally unreachable there (unless the program itself calls DoEvents). The receipt carries diag.dispatch{toolEntryToObserveMs, addInObserveMs, queuedBeforeToolEntryMs, dispatchedLate, programStateAtEntry}; when the call was dispatched AFTER the program ended, dispatchedLate is true and message says so. For real run-time observation use a breakpoint + run_until_break (BREAK state IS served), or make the program call DoEvents.")

    Dim pLe As Object
    Set pLe = CreateObject("Scripting.Dictionary")
    Set pLe("dismiss") = Prop("boolean", "true = dismiss the runtime-error box by clicking its End button with send_window_message (the ONLY working channel: close_active_dialogs refuses it by the safe whitelist - measured). Nothing is clicked when false.")
    tools.Add MakeTool("last_error", _
        "Return the UNHANDLED-ERROR context WITHOUT evaluating in the break frame: {errorDialog{present,title,body,buttons[]}, errNumber/errDescription/errSource + errParsed{...}, callstack, localsNames[] (the names an earlier get_locals enumerated), position, proc, module, stopped_at, howToRecover}. " & _
        "SOURCE DECLARATION: the error text comes from the runtime-error dialog BODY (Win32 window channel), the stack from the VBA6 callstack probe, and the NAME list from get_locals' source-declaration scan; NOTHING is read by break-frame evaluation. Fields that cannot be parsed are returned as null with parsed:false - never guessed. " & _
        "The box is MODAL: dismiss it with dismiss:true (send_window_message click on the End button) or the IDE stays blocked." & _
        "RE-1 (v1.16.3, ADDITIVE): the receipt ALSO carries a STRUCTURED error block read from the VBA6 exception record by pure in-process memory reads (GetModuleHandleW + CopyMemory; NO engine call): err{number,description,source,hresult,numberBasis}, errBasis (vb6map | dialog | none | unavailable), errChannel{readOk,recordUsed,selector,fields,bstrReads,why}, errCrossCheck{available,number,description,relation,note} and errLocation{line,column,file,basis,why}. This ADDS fields WHILE THE RUNTIME-ERROR BOX IS STILL THERE: it does not suppress the box, does not auto-dismiss it, and does not read errors without a box - automatic dismissal and box-less reading are NOT part of this batch." & _
        " It also removes the old dependence on the localized dialog TEXT for number/description/source (the dialog body is still read and still feeds the unchanged errNumber/errDescription/errSource/errParsed keys). Unreadable fields are null with errBasis/errChannel.why naming the reason - never guessed, and errLocation is always null with basis=not_attempted (no line/column is invented). errChannel.selector.basis=unavailable means the SELECTOR DID NOT PARTICIPATE in choosing the record (record A is the live record, measured in BOTH the modal and the settled state) - it is NOT a read failure, and no new enum value is added.", _
        SchemaWith(pLe, Req()))

    Dim pIt As Object
    Set pIt = CreateObject("Scripting.Dictionary")
    Set pIt("name") = Prop("string", "Component name")
    Set pIt("procedure") = Prop("string", "Procedure to instrument")
    Set pIt("expressions") = Prop("string", "Comma-separated expressions, e.g. x,y")
    Set pIt("id") = Prop("string", "Trace id (also the removal handle)")
    Set pIt("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("inject_trace", _
        "Insert a McpTrace call at the start of a procedure. The guard is a RUNTIME 'If McpTraceOn() Then' (not '#If MCP_TRACE'), because a compile-time condition only applies when the project is loaded with that CondComp - so this needs no CondComp and no reopen; inject_trace now GENERATES AND ATTACHES the module modMcpTrace itself when the project does not have it (nothing to import by hand) and verifies the attachment before answering; the generated header comment language is chosen UI-FIRST (the IDE's own UI language when it can be read, else the VB6.EXE version resource, else English) and the receipt reports all three signals. The sink is enabled by the MCP_TRACE_ON / MCP_TRACE_FILE environment variables of the VB6 process (see the response hints).", _
        SchemaWith(pIt, Req("name", "procedure", "id")))

    Dim pRt As Object
    Set pRt = CreateObject("Scripting.Dictionary")
    Set pRt("name") = Prop("string", "Component name")
    Set pRt("id") = Prop("string", "Trace id to remove")
    Set pRt("expected_sha256") = Prop("string", "Optional guard")
    tools.Add MakeTool("remove_trace", _
        "Remove a previously injected trace block; sha256 returns to the pre-injection value. The modMcpTrace component is NEVER removed automatically - the receipt always reports moduleRemoved=false, moduleLeftInPlace=true and moduleRemoveReason=""kept_by_design"" (moduleRemoveReason is module_absent when no such component exists), because (a) you may have added your own code to it and (b) other @mcp-trace blocks elsewhere may still call McpTrace/McpTraceOn, so the project would stop compiling without it. moduleRemoveNote carries the manual removal steps (verify no '@mcp-trace' block remains, then call remove_component yourself). The old automatic removal (removed_autogen_module) no longer exists.", _
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

    ' ---- R2 + R12: live desktop / input-channel probe (read-only) ----
    tools.Add MakeTool("desktop_probe", _
        "Read the LIVE desktop/interactivity state of the process that hosts the add-in: {ok, windowStation, desktop, sessionId, interactive, canSendInput, foregroundHwnd, foregroundReadable, probeBasis, note}. Every value is MEASURED on each call (GetProcessWindowStation + GetUserObjectInformationW(UOI_NAME) + GetThreadDesktop + OpenInputDesktop + GetForegroundWindow + IsWindow + ProcessIdToSessionId, all inside the add-in process; the input-desktop handle it opens is closed again). The foreground/desktop state is reported as a LIVE structured value: foregroundHwnd can be 0 on a non-interactive or locked desktop and that is reported as-is. The probe never steals focu" & _
        "s to obtain it (no SetForegroundWindow) and never derives send-input capability from a window-station name: this add-in delivers input as PostMessage(WM_KEYDOWN/WM_KEYUP/WM_MOUSEWHEEL) to windows of its OWN process, measured to deliver while ANOTHER process owns the foreground. If any of those APIs cannot be called, or returns no handle, the tool answers ok:false + failureKind=api_unavailable + failedApi + apiError + apiErrorText + measured:false instead of a zero-filled picture, and the desktop{} block on the input-family receipts carries measured:false as well. A missing Alias clause on a Declare surfaces here as VB error 453 rather than as zeros. Read-only.", _
        SchemaWith(Nothing))

    Dim pVp As Object
    Set pVp = CreateObject("Scripting.Dictionary")
    tools.Add MakeTool("vba6_probe_plan", _
        "Ordered plan for the VBA6.DLL debug-engine probes (read-only first, writes last) with preconditions. Run only inside the IDE process.", _
        SchemaWith(pVp))

    Dim pEa As Object
    Set pEa = CreateObject("Scripting.Dictionary")
    Set pEa("name") = Prop("string", "Component name. OMITTED or EMPTY = bounded whole-project export: every component is parsed under a wall-clock budget and the receipt carries per-component counts + budgetExhausted/elapsedMs, instead of scanning 110 components unbounded.")
    Set pEa("membersOnly") = Prop("boolean", "Light mode (default false): return only the member list (kind/name/sig) and skip the TSV spec text, which also skips a second parse of the same text. This tool never compares anything, so membersOnly is OUTPUT SIZE only (no comparison is skipped). Unknown argument names are refused with reason=unknown_argument + acceptedArguments.")
    tools.Add MakeTool("export_public_api", _
        "Freeze the component public API as a spec (JSON members[] plus a TSV spec field). Text level; parameter names are dropped. With no name, exports every component of the project under a wall-clock budget (partial result + budgetExhausted/elapsedMs); membersOnly=true skips the spec text.", _
        SchemaWith(pEa))

    Dim pCa As Object
    Set pCa = CreateObject("Scripting.Dictionary")
    Set pCa("name") = Prop("string", "Component name")
    Set pCa("baseline") = Prop("string", "Baseline spec TSV previously returned by export_public_api (its spec field)")
    ' (the schema-enum finding): the implementation at ToolCheckPublicApi REQUIRES a non-empty
    ' baseline (Err.Raise "baseline spec is required"), and 's central classifier already
    ' answers that condition as reason:"required_argument_missing" / argument:"baseline". The schema,
    ' however, declared NO "required" key at all, so a schema-validating caller could omit baseline
    ' and only learn about it from a runtime refusal. DIRECTION CHOSEN: DECLARE it (below).
    ' Rejected alternative - "accept an empty baseline": it would have to INVENT a success (every
    ' current member reported as added[]), i.e. exactly the misleading-success class this product
    ' treats as a defect (the dry-run findings), and it would delete the classification delivered.
    ' NO BEHAVIOUR CHANGE: a missing baseline was refused before and is refused now; only the
    ' DECLARATION moves (see 内部证据 section (1)).
    Set pCa("membersOnly") = Prop("boolean", "Output size only (default false): the comparison against baseline ALWAYS happens. membersOnly=true ADDS the member list (kind/name/sig) to the receipt, so one call returns the member table AND violations[]/added[]/counts; it is never a switch that skips the comparison.")
    tools.Add MakeTool("check_public_api", _
        "Compare the current component public API against a frozen spec; returns violations[] (removed / changed), added[] and counts{}. baseline is REQUIRED and the comparison ALWAYS runs when a component name is given - no combination of arguments skips it, and a baseline that is not spec text is refused with reason=baseline_unusable instead of being reported as 'no differences'. With no name, checks every component of the project under a wall-clock budget (per-component counts, partial result + budgetExhausted/elapsedMs). membersOnly=true is OUTPUT SIZE only: the comparison still runs and the member list is added to the receipt. Unknown argument names are refused with reason=unknown_argument + acceptedArguments. LIMITS of membersOnly: the member table is a LIGHT text parse (kind / name / signature, parameter NAMES dropped by design) with no type resolution - it is not a compiler symbol table, and because the comparison always runs it cannot be used to skip one.", _
        SchemaWith(pCa, Req("baseline")))

    Dim pDg As Object
    Set pDg = CreateObject("Scripting.Dictionary")
    Set pDg("char_quota") = Prop("integer", "Optional character quota for THIS call (overrides VB6MCP_CHAR_QUOTA and the default). With budget_ms this is how a caller requests a COMPLETE graph for a large project; the receipt reports edgesComplete + componentsUnscanned/edgeComponentsUnscanned so completeness is decidable, never guessed.")
    Set pDg("budget_ms") = Prop("integer", "Optional WALL-CLOCK budget for THIS call (overrides VB6MCP_BUDGET_MS and the 30 s default). The wall clock is a separate last-resort cutoff reported as timingTruncated; a large project needs both this and char_quota raised for a complete graph (measured: about 35 s for a 41,637-line project).")
    tools.Add MakeTool("dependency_graph", _
        "Text level dependency graph of the active project: nodes, edges (component name references), cycles, and a purity classification per component: plus a PROCEDURE-level caller/callee graph (procNodes / procEdges / procCount). v1.11.3: the build is bounded by a DETERMINISTIC character quota (env VB6MCP_CHAR_QUOTA, default 900000), so the SAME input always returns the SAME receipt; every phase keeps a reserved share of that quota, so a default run reports nodes AND some procedure edges. Large projects are PARTIALLY covered by default - read partial, budgetExhausted (the deterministic quota fired) vs timingTruncated (the wall clock fired - the LAST resort, and NOT reproducible), componentsScanned/componentsTotal, procEdgesFound and the *StageCompleted flags (a stage that was cut reports false; a stage whose input was itself cut reports false too). Raise VB6MCP_CHAR_QUOTA (and VB6MCP_BUDGET_MS, since the clock is what would stop it first) for a complete graph; a complete run is reproducible.", _
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
    Set pVb("exec_probe") = Prop("boolean", "Run the Immediate-window executor probe (EbExecuteLine) instead of the standard set. This export (ordinal 147, RVA 0x127065) IS the Immediate-window line executor.")
    Set pVb("exec_line") = Prop("string", "Line to execute, e.g. 1+1 or Debug.Print ... (Immediate-window syntax)")
    Set pVb("exec_line_file") = Prop("string", "Path to a UTF-8/ANSI text file whose text is used as exec_line. Mutually exclusive with exec_line; both given => reason both_inline_and_file; a missing file => reason input_file_not_found with path; a file larger than 65536 bytes => reason input_file_too_large. On success the receipt echoes inputFile/inputBytes/inputEncoding.")
    Set pVb("exec_a2") = Prop("integer", "arg2 (semantics unresolved; values 0/1 are the candidates)")
    Set pVb("exec_a3") = Prop("integer", "arg3 (unresolved)")
    Set pVb("exec_a4") = Prop("integer", "arg4 (mode flag: the epilogue uses setz al, so 0 vs 1 differ)")
    Set pVb("exec_scope") = Prop("boolean", "Set the eval-scope bit (WORD[VBA6+0x17D076] |= 0x100) around the call, the way the IDE's own immediate-window path does, then restore it. This is the cheapest test of whether locals become visible at a breakpoint.")
    Set pVb("exec_dump_ctx") = Prop("boolean", "Pass arg2 a 64-byte buffer and dump the first 32 bytes afterwards. WARNING: measured to crash the IDE (AV in VBA6.dll) - arg2 is dereferenced as a structured pointer, so leave it 0 unless the object type is known.")
    Set pVb("exec_set_frame") = Prop("boolean", "Make the breakpoint's frame the current evaluation frame (write dword_global VBA6+0x17D0A4 from the engine+0x2F0 chain's first kind==1 node payload) around the call, then restore. Expected to make locals visible at a breakpoint.")
    Set pVb("frame_probe") = Prop("boolean", "Read-only two-state observation - EbMode, dword_global VBA6+0x17D0A4 (evaluation frame), and the chain at engine+0x2F0 with each node's kind/payload. Writes nothing.")
    Set pVb("frame_set") = Prop("boolean", "Call the engine's own frame setter sub_FBC3163 (RVA 0x133163) through an executable thunk, with frame_set_a1/frame_set_a2 (default 0/1 = find the named frame). This is the authoritative way to bind the evaluation frame.")
    Set pVb("frame_set_a1") = Prop("integer", "frame_set arg1: frame index (default 0)")
    Set pVb("frame_set_a2") = Prop("integer", "frame_set arg2: 1 = walk to the first kind==1 frame (default 1)")
    Set pVb("rva_call") = Prop("string", "Call an INTERNAL function by RVA; optional vb6: prefix targets the HOST image VB6.EXE instead of VBA6.DLL. Call a VBA6 function by RVA (hex, e.g. 0x23EA) through an executable thunk, with rva_argc/rva_a1/rva_a2. Unsafe by nature: the callee must be __stdcall with argc args and clean its own stack.")
    Set pVb("rva_argc") = Prop("integer", "Argument count 0..2 (must match the callee's ret N)")
    Set pVb("rva_a1") = Prop("integer", "Argument 1")
    Set pVb("rva_a2") = Prop("integer", "Argument 2")
    Set pVb("worker_call") = Prop("boolean", "DECISIVE: call the real worker sub_FB7317A (RVA 0xE317A) the way the IDE does - __thiscall ecx=engine, 8 stack args, retn 20h. Can crash the IDE if an argument is wrong.")
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
    Set pVb("worker_p2_engine") = Prop("boolean", "Pass p2 = the engine singleton instead of the IDE's literal 1. NOTE: with p1=0 the worker may dereference p2 as an object, and 1 then crashes.")
    Set pVb("mem_dump") = Prop("boolean", "READ-ONLY memory dump (locate the ExFrame stack top / check frame offsets). Safe: refuses addresses outside user space.")
    Set pVb("mem_rva") = Prop("string", "Address as an RVA (hex, e.g. 0x17D000). Default image is VBA6.DLL; prefix vb6+ to address the HOST image VB6.EXE (e.g. vb6+0x1234) or vba6+ to be explicit. Used when mem_addr is 0")
    Set pVb("mem_addr") = Prop("string", "Absolute address to read (hex, e.g. 0x01234567); takes precedence over mem_rva")
    ' later additions (probe-only; the two are independent of each other)
    Set pVb("hosted_expr_srv") = Prop("boolean", "NEW DIRECTION: instantiate VBA6's HOSTED expression service via CreateHostedExprSrv (one of the 317 exports the IDE never imports), then report the object pointer it returns and the first 10 slots of its vtable. Read-mostly: it allocates a 404-byte engine object; nothing is patched.")
    Set pVb("wnd_probe") = Prop("boolean", "READ-ONLY, in-process: enumerate the DIRECT CHILDREN (depth exactly 2) of this process's TOP-LEVEL windows and classify both the window procedure and the module that registered the window class (vb6 / vba6 / other). The top-level windows themselves are NEVER listed and grandchildren or deeper are NEVER listed; there are no visibility, owner, rect or style filters (source: Vba6Probe.bas WindowProbeJson). Cross-process GetWindowLong/GCL_HMODULE both return 0, so this probe only works from inside the IDE. WARNING: this is NOT a complete window enumeration - it saw only 27 of the 76 windows that find_gui returned (about 36 percent), so any does-not-exist conclusion about a window or a dialog drawn from wnd_probe is invalid. STRUCTURAL BLIND SPOT: a dialog is normally a TOP-LEVEL #32770 window, which wnd_probe structurally cannot see - use find_gui (all top-level windows plus the whole recursive subtree) to find dialogs or any complete set.")
    Set pVb("wnd_max") = Prop("integer", "Maximum windows to report for wnd_probe (default 200, cap 1500)")
    Set pVb("mem_scan") = Prop("boolean", "READ-ONLY byte-wise dword scanner over a live range: reports every offset whose dword lies in [scan_lo, scan_hi], with its alignment. Used to replicate the whole-image scan of VB6.EXE for absolute references into the VBA6 engine globals.")
    Set pVb("scan_addr") = Prop("string", "Start address for mem_scan (hex, e.g. 0x400000)")
    Set pVb("scan_len") = Prop("integer", "Bytes to scan for mem_scan (max 1048576)")
    Set pVb("scan_lo") = Prop("integer", "Low bound of the dword value range to match (must be non-zero)")
    Set pVb("scan_hi") = Prop("integer", "High bound of the dword value range to match")
    Set pVb("scan_max") = Prop("integer", "Maximum hits to return (default 64, cap 512)")
    Set pVb("export_rvas") = Prop("boolean", "READ-ONLY: resolve the exports the RE names BY NAME through the real PE export table and report each one's true RVA, so any 'function X is at RVA 0xNNNN' claim can be checked in one call.")
    Set pVb("engine_items") = Prop("boolean", _
        "READ-ONLY: calls VBA6 EbGetStatusOfItems(out, size) the way VB6.EXE's command dispatcher does. " & _
        "The callee ZEROES a 72-byte (0x48) local block (18 dword slots), fills status bits, then copies exactly size bytes to the caller. " & _
        "This is a STATUS BITMAP, NOT a list of 72 capability items, and item-index semantics are NOT supported. " & _
        "Receipt: bitmap_len (72), bitmap_hex (144 lowercase hex chars), byte_offsets[] for the bits inferred from static write-site analysis, " & _
        "and items/available both null (historical names, kept only to state explicitly that item numbers are meaningless). No writes.")
    Set pVb("cs_fn") = Prop("integer", "Callback-free single-frame probe: call EbGetCallstackFunction(frame, &outA, &outB, &outC, &arg5) with the authoritative IDE prototype (3 out-BSTRs zeroed, arg5 a zeroed writable dword, rc=0 means success). frame must be 0..count-1. Omit or -1 to skip.")
    Set pVb("mem_len") = Prop("integer", "Bytes to read, multiple of 4, max 512 (default 64)")
    ' Declared so an existing caller keeps working now that undeclared keys are refused:
    ' the probe body does not read them in this build, and the receipt reports the engine's
    ' OWN stepMode / callstack values instead of an echo of these arguments.
    Set pVb("step") = Prop("string", "Accepted for compatibility with the probe plan; this build's probe body does not read it (the receipt reports the engine's own stepMode)")
    Set pVb("frame_index") = Prop("integer", "Accepted for compatibility with the probe plan; this build's callstack probe reports the engine's own frames (use cs_fn for one explicit frame)")
    ' ---- project change journal (IDE event sinks) ----
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
        "Read or write the .vbp version resource and build metadata (MajorVer/MinorVer/RevisionVer/AutoIncrementVer/CompanyName/FileDescription/ProductName/legal notices/Comments/Title/ExeName32/CondComp). These keys are not exposed by the VBIDE object model, so they are edited as text: component and reference lines are never touched and a .bak is kept. VB project/source files (.vbp/.frm/.cls/.bas/.ctl) are written ANSI without BOM; never hand-write them as UTF-8-with-BOM (VB6 reads a mangled first keyword and the project stops loading).", _
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
    Set pVs("baseline") = Prop("string", "Optional lint baseline (a previous verify/check_component receipt text or its stableKey lines). Diagnostics matching the baseline are filtered, so pre-existing lint noise cannot inflate the numbers; the pass/fail/unknown counters keep pass+fail+unknown==checked.")
    Set pVs("requireCompiler") = Prop("boolean", "false (default) keeps this receipt byte-for-byte. true obtains ONE project-level COMPILER verdict (an isolated VB6.EXE /make over a COPY of the project on disk) before the sweep, and REFUSES structurally when no compiler is available (ok:false + reason=compiler_not_configured + compilerProbe{resolved,path,why} + howToProceed) instead of letting lint stand in for the compiler. The pass+fail+unknown==checked invariant is unchanged; on success the receipt adds authority=compiler + compiler{path,logPath,errorCount,artifactChanged}.")
    tools.Add MakeTool("verify_scope", _
        "Read-only: for a set of components report {component, ideEqualsDisk, lintErrors, lintWarnings, firstDiagnostic} with pass/fail/unknown counters; only failing (and unknown) components are listed in detail. COUNT SEMANTICS (v1.11.9): the three counters are mutually exclusive - pass = verify answered PASS, fail = verify answered FAIL (a PROVEN inconsistency), unknown = the verdict could not be proven either way (e.g. only the build evidence is missing) and is NOT a fail. SCALE OF THE BUILD EVIDENCE: it is PROJECT-level - buildOk/buildSource/buildNote are read from a project-wide build cache, so no single component can be the only one missing it (on a never-built project every consistent component is UNKNOWN and pass is unreachable). INVARIANT: pass + fail + unknown == checked on every receipt, failures[] holds proven inconsistencies only, unknowns[] holds the unknowns only, and no component appears in both; countsConsistent publishes the check. " & _
        "requireCompiler:true obtains ONE project-level COMPILER verdict before the sweep and refuses with reason=compiler_not_configured instead of letting lint stand in for the compiler; the counters and countsConsistent keep their documented meaning.", _
        SchemaWith(pVs))

    tools.Add MakeTool("todo_status", _
        "READ-ONLY task-board snapshot. Scans an OPTIONAL in-repo task folder (tasks/ + reports/ Markdown) and returns {taskCount, openTasks, tasks[], reportCount, reports[], readOnly}. A missing task folder is a normal result (0 tasks), not an error.", _
        SchemaWith(Nothing))

    ' ---- contract v1.3.0: machine-readable version identity ----
    tools.Add MakeTool("mcp_info", _
        "Read-only version stamp for this server: {ok, product, contract, toolCount, build:{sha256, builtAt}, breaking[], migrations[], changelog[]}. build.sha256 is the SHA-256 of the DLL file on disk, read at call time (never a compile-time literal), so an independent Get-FileHash can be compared with it; builtAt is that file's LastWriteTime. breaking[] lists the contract changes that alter how existing calls must be made. Call this once after initialize to learn which contract the server implements.", _
        SchemaWith(Nothing))

    ' ---- contract v1.6.0: cross-component replace + component undo ----
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

    ' ---- contract v1.9.0 (A B): foreground safety + immediate execution ----
    Dim pImE As Object
    Set pImE = CreateObject("Scripting.Dictionary")
    Set pImE("expression") = Prop("string", "Expression to evaluate in the ACTIVE project's immediate context, e.g. 1/4 or UCase$(""ab""). It is evaluated exactly ONCE.")
    Set pImE("expression_file") = Prop("string", "Path to a UTF-8/ANSI text file whose text is used as expression. Mutually exclusive with expression; both given => reason both_inline_and_file; a missing file => reason input_file_not_found with path; a file larger than 65536 bytes => reason input_file_too_large. On success the receipt echoes inputFile/inputBytes/inputEncoding.")
    Set pImE("timeout_ms") = Prop("integer", "Gate-age threshold in ms; if this call never returns, the ARMED release frees the single-flight gate on the next tool call (default 5000). It is NOT a hard timeout - the add-in cannot preempt a statement.")
    tools.Add MakeTool("immediate_eval", _
        "Evaluate a SCALAR EXPRESSION in the VB6 IDE immediate context and return its VALUE: {ok, value, valueKind, typeName, output, elapsedMs}. " & _
        "Object/array EXPRESSIONS (CreateObject(/GetObject(/New/Array() are REFUSED BEFORE EXECUTION with reason=object_expression_unsupported: the immediate channel does not support them (measured: a plain object assignment hangs the call). " & _
        "Assign the object to a PROJECT variable first and read scalar members; scalar results are unchanged. The value is carried out through a private temp file the tool reads back, so no Immediate-window pane read is involved. " & _
        "An unparseable sink is reported as {ok:false, reason:""value_not_parsed""} together with the RAW output - nothing is guessed. A statement that RAISES still blocks the IDE (VB6 refuses On Error on this channel with ret=40036, measured), so timeout_ms is a gate-age threshold rather than a hard timeout; ide_state{release_inflight:true} forces the gate open afterwards. The expression may also be supplied through expression_file (mutually exclusive with expression).", _
        SchemaWith(pImE, Req("expression")), _
        "Sink failures are typed: a rejected expression is reason=""engine_refused_line""; a SINK line failure (file number/path/permission) is reason=""eval_sink_error"" with errNumber/errDescription/lineText - never reported as an expression refusal.")

    Dim pImR As Object
    Set pImR = CreateObject("Scripting.Dictionary")
    Set pImR("code") = Prop("string", "One or more statements to run, e.g. Debug.Print 1+1 or x = 5: Debug.Print x. A CR/LF in the input is joined with a colon and a space; each statement-boundary Debug.Print is redirected into the sink.")
    Set pImR("code_file") = Prop("string", "Path to a UTF-8/ANSI text file whose text is used as code. Mutually exclusive with code; both given => reason both_inline_and_file; a missing file => reason input_file_not_found with path; a file larger than 65536 bytes => reason input_file_too_large. On success the receipt echoes inputFile/inputBytes/inputEncoding.")
    Set pImR("timeout_ms") = Prop("integer", "Gate-age threshold in ms; if this call never returns, the ARMED release frees the single-flight gate on the next tool call (default 5000). It is NOT a hard timeout.")
    tools.Add MakeTool("immediate_run", _
        "RUN VB statements in the VB6 IDE's own immediate context and return what they PRINTED, with the caller NOT having to write Open/Print/Close: {ok, output, outputTrimmed, outputLines, debugPrintRewrites, elapsedMs}. Each statement-boundary Debug.Print is redirected into a private temp file, and only the text produced by THIS call is returned. The ? shorthand and _ line continuations are not supported. A statement that RAISES still blocks the IDE - see immediate_eval; timeout_ms arms a self-healing release and ide_state{release_inflight:true} forces the gate open. The statement body may also be supplied through code_file (mutually exclusive with code).", _
        SchemaWith(pImR, Req("code")), _
        "Sink failures are typed: a rejected statement is reason=""engine_refused_line""; a SINK line failure (file number/path/permission) is reason=""run_sink_error"" with errNumber/errDescription/lineText - never reported as a statement refusal.")

    Dim pQm As Object
    Set pQm = CreateObject("Scripting.Dictionary")
    Set pQm("on") = Prop("boolean", "true = turn quiet mode ON (the add-in stops taking the user's foreground and every receipt carries a quiet{} block with an advisory); false = turn it OFF and restore the previous behaviour byte for byte.")
    Set pQm("ttl_ms") = Prop("integer", "Optional lifetime in ms (default 600000 = 10 minutes). The mode expires by itself once the TTL passes; the TTL is re-evaluated on every read.")
    tools.Add MakeTool("ide_quiet_mode", _
        "Turn the SESSION-LEVEL quiet mode on or off. While ON the add-in does not call Window.SetFocus for control_window{action:""activate""} (that receipt reports quietModeSuppressedFocus:true plus the action that was still performed) and it gives the user's foreground back after a menu-driven action (native_breakpoint); EVERY tools/call receipt then carries quiet{quietMode,since,ttlMs,quietExpiresAt,suppressedFocusCount,advisory}. The advisory asks the human not to type into this IDE" & _
        " while an agent drives it, because keystrokes aimed at another window can land in the IDE when it is raised. quiet OFF is byte-for-byte the previous behaviour. The foreground/desktop state is reported as a LIVE structured value (see desktop_probe): the handle can be 0 on a non-interactive or locked desktop and that is reported as-is, so foregroundReadable=false means the handle could not be read - never read that as ""the focus was left alone"". This tool is read-only: it changes only the add-in's own state, so it is still served while a write is in flight.", _
        SchemaWith(pQm, Req("on")))

    Set BuildToolsListResult = root
    If cacheOff = 1 Then Set cacheRoot = root
End Function

Public Function MakeTool(ByVal name As String, ByVal desc As String, ByVal schema As Object, Optional ByVal note As String = "") As Object
    Dim t As Object
    Set t = CreateObject("Scripting.Dictionary")
    t("name") = name
    If Len(note) > 0 Then
        If Left$(note, 1) = " " Then
            t("description") = desc & note
        Else
            t("description") = desc & " " & note
        End If
    Else
        t("description") = desc
    End If
    Set t("inputSchema") = schema
    Set MakeTool = t
End Function

Public Function SchemaWith(ByVal props As Object, Optional ByVal required As Object) As Object
    ' JSON Schema shape is NORMALISED here, once, for every tool:
    '   properties : ALWAYS an object (an empty Dictionary when there are no parameters)
    '   required   : ALWAYS an ARRAY (an empty Collection when nothing is required)
    ' A caller that passes a Dictionary by mistake is coerced to its KEYS, so a
    ' malformed "required": {} can never reach the wire again (a JSON-Schema
    ' validator is entitled to reject a non-array required).
    Dim s As Object
    Dim reqList As Object
    Dim one As Variant
    Dim k As Variant
    Set s = CreateObject("Scripting.Dictionary")
    s("type") = "object"
    If props Is Nothing Then
        Set s("properties") = CreateObject("Scripting.Dictionary")
    Else
        Set s("properties") = props
    End If
    Set reqList = New Collection
    If Not required Is Nothing Then
        If TypeOf required Is Collection Then
            For Each one In required
                reqList.Add CStr(one)
            Next one
        Else
            ' vbObject: accept any object that exposes Keys (Scripting.Dictionary)
            On Error Resume Next
            For Each k In required.Keys
                reqList.Add CStr(k)
            Next k
            On Error GoTo 0
        End If
    End If
    Set s("required") = reqList
    Set SchemaWith = s
End Function

Public Function Prop(ByVal t As String, ByVal desc As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("type") = t
    d("description") = desc
    Set Prop = d
End Function

Public Function EnumList(ParamArray values() As Variant) As Object
    Dim c As Object
    Set c = New Collection
    Dim i As Long
    For i = LBound(values) To UBound(values)
        c.Add CStr(values(i))
    Next
    Set EnumList = c
End Function

Public Function Req(ParamArray names() As Variant) As Object
    Dim c As Object
    Set c = New Collection
    Dim i As Long
    For i = LBound(names) To UBound(names)
        c.Add CStr(names(i))
    Next
    Set Req = c
End Function

Public Function PropValidate() As Object
    Dim p As Object
    Set p = Prop("string", "off = no lint, warn (default) = report diagnostics and still write, strict = refuse the write when a severity=error diagnostic exists")
    Set p("enum") = EnumList("off", "warn", "strict")
    Set PropValidate = p
End Function
