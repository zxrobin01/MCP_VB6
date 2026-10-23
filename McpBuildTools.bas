Attribute VB_Name = "McpBuildTools"
Option Explicit

' =======================================================================
' C-3 facade batch: McpBuildTools (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Private Declare Function WoGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Public Function BuildAuthorityNote(Owner As CodeTools, ByVal bSource As String, ByVal compiledFrom As String) As String
    Select Case LCase$(Trim$(bSource))
        Case "cache-fresh"
            BuildAuthorityNote = "the build cache record is current for this project (fingerprint and hash match), so its verdict speaks for the source it compiled (buildCompiledFrom=" & compiledFrom & ")"
        Case "cache-pane-disk-mismatch"
            BuildAuthorityNote = "the recorded build compiled the IDE PANE while the DISK differed; its verdict does NOT describe the disk file this verify compares against - save the project and rebuild"
        Case "cache-stale"
            BuildAuthorityNote = "the project changed after the recorded build; no current build evidence exists for the disk source"
        Case "cache-other-project"
            BuildAuthorityNote = "the build cache holds a different project; nothing is claimed about this one"
        Case "log"
            BuildAuthorityNote = "the verdict came from a compiler log passed to verify"
        Case Else
            BuildAuthorityNote = "no build evidence: no compiler ran for this project in this session"
    End Select
End Function

Public Function BuildAuthorityOf(Owner As CodeTools, ByVal bSource As String) As String
    ' FIX-27 (R4.7): a CLOSED set of authority kinds for verify's build evidence.
    Select Case LCase$(Trim$(bSource))
        Case "cache-fresh"
            BuildAuthorityOf = "cache_current"
        Case "cache-pane-disk-mismatch"
            BuildAuthorityOf = "pane_compiled_cache"
        Case "cache-stale"
            BuildAuthorityOf = "cache_stale"
        Case "cache-other-project"
            BuildAuthorityOf = "other_project_cache"
        Case "log"
            BuildAuthorityOf = "compiler_log"
        Case "none", ""
            BuildAuthorityOf = "none"
        Case Else
            BuildAuthorityOf = LCase$(Trim$(bSource))
    End Select
End Function

Public Function BuildBuiltAt(Owner As CodeTools) As String
    If Owner.m_buildBuiltAtDone Then
        BuildBuiltAt = Owner.m_buildBuiltAt
        Exit Function
    End If
    Owner.m_buildBuiltAtDone = True
    On Error Resume Next
    Dim p2 As String
    p2 = Owner.OwnDllPath()
    If Len(p2) > 0 Then Owner.m_buildBuiltAt = Format$(FileDateTime(p2), "yyyy-mm-dd hh:nn:ss")
    Err.Clear
    BuildBuiltAt = Owner.m_buildBuiltAt
End Function

Public Function BuildInstructions(Owner As CodeTools) As String
    Dim s As String
    s = "PRODUCT " & VB6MCP_PRODUCT & " " & VB6MCP_PRODUCT_VERSION & "  CONTRACT " & VB6MCP_CONTRACT_VERSION & "  TOOLS " & CStr(Owner.ToolCount()) & "  BUILT " & Left$(BuildSha256(Owner), 8) & vbCrLf
    s = s & "HISTORICAL CONTRACT NOTES (oldest first; the authoritative breaking[] is in mcp_info):" & vbCrLf
    s = s & "- get_component_code: coords default disk -> ide (pass coords=""disk"" for the old text/numbering); response is a JSON envelope, text in result.text" & vbCrLf
    s = s & "- replace_code: out-of-range/continuation now REJECTED (range_beyond_end / inside_continuation); deletedLines=0 is no longer a success signal" & vbCrLf
    s = s & "- write-class tools: validate=off|warn|strict; edit-side failures carry ok:false + reason" & vbCrLf
    s = s & "- write_file now truncates; build/make_project/verify: judge by artifactChanged/verdict (compiled-no-artifact-change = NOT built)" & vbCrLf
    ' --- usage guidance: routing + encoding (guidance audit) ------------------
    s = s & "ROUTING (use the tool; never hand-edit VB project files):" & vbCrLf
    s = s & "- .vbp version + build metadata (MajorVer / MinorVer / RevisionVer / AutoIncrementVer / CompanyName / FileDescription / ProductName / Title / ExeName32 / CondComp) -> project_meta (text edit, keeps a .bak; close the project first)" & vbCrLf
    s = s & "- project / IDE options (description / help file / startup object / startup project / output path / read-only) -> set_project_options" & vbCrLf
    s = s & "- UAC level + DPI awareness -> app_manifest (writes <exe>.manifest; do not hand-write one)" & vbCrLf
    s = s & "- components: add / remove / import / rename / reload / export -> add_component / remove_component / import_component / rename_component / reload_component / export_component" & vbCrLf
    s = s & "- COM references -> add_reference / remove_reference / check_references" & vbCrLf
    s = s & "- forms and controls: list / add / remove / design-time properties -> form_designer / add_control / remove_control / set_control_property / list_component_parts" & vbCrLf
    s = s & "- lint before writing / static checks / formatting -> validate_code / check_component / format_code / verify / verify_scope" & vbCrLf
    s = s & "- component text edits (surgical or bulk) -> edit_by_anchor / apply_edits / replace_code / replace_in_files; preview_edit first; snapshot=true or undo_edit to roll back" & vbCrLf
    s = s & "- project-wide checkpoint / rollback -> project_snapshot (whole project: .vbp + component files)" & vbCrLf
    s = s & "- build verdict -> build (offline /make) or make_project (in-IDE): judge by artifactChanged + verdict; compiled:true alone is NOT built" & vbCrLf
    s = s & "- ENCODING: VB project files (.vbp / .frm / .cls / .bas / .ctl / .pag / .dsr / .res / .inc) must stay system ANSI (GBK) WITHOUT a BOM - never hand-edit or re-encode them (a UTF-8 BOM makes VB6 refuse the project: ""invalid keyword""); the tools above already write ANSI for those extensions." & vbCrLf
    s = s & "- " & VB6MCP_CONTRACT_VERSION & " current (" & CStr(Owner.ToolCount()) & " add-in tools): the bullets above are historical - call mcp_info once after initialize for the authoritative breaking[] / migrations[] / changelog[]" & vbCrLf
    s = s & "- v1.7.0 snapshot param (ADDITIVE, non-breaking): replace_code / apply_edits / edit_by_anchor accept snapshot=true (on/yes/1) to guarantee exactly ONE pre-edit pane snapshot that undo_edit can restore; default off (set / apply_edits / edit_by_anchor already snapshot by default)" & vbCrLf
    s = s & "CAPABILITY INDEX (question -> tool; read this before improvising):" & vbCrLf
    s = s & "- impact / who calls whom -> dependency_graph  (component edges + PROCEDURE-level procNodes/procEdges/procCount + cycles + purity)" & vbCrLf
    s = s & "- who references a symbol -> rename_symbol  (dry_run=true by default; skips strings and comments)" & vbCrLf
    s = s & "- public API broken? -> check_public_api / export_public_api ; missing files or OCX parts -> check_references / list_component_parts" & vbCrLf
    s = s & "- runtime window tree / control text / modal dialog -> find_gui / ui_text / read_active_dialogs" & vbCrLf
    s = s & "- click a button or send keys to a window of THIS VB6 process -> send_window_message" & vbCrLf
    s = s & "- offline single-module test -> test_host ; stub a procedure or trace calls -> fault_inject / inject_trace" & vbCrLf
    s = s & "- NOT available: screenshots; driving a standalone built EXE (message tools are in-process only)" & vbCrLf
    s = s & "FILE ENCODING: .vbp/.frm/.cls/.bas/.ctl are ANSI WITHOUT BOM (write_file / write_source already do this); never write them as UTF-8-with-BOM, or VB6 reads a mangled first keyword and the project stops loading; use project_meta for .vbp version metadata (keeps a .bak)." & vbCrLf
    BuildInstructions = s
End Function

Public Function BuildOnDiskChanged(Owner As CodeTools) As Boolean
    On Error Resume Next
    Dim now As Long
    now = WoGetTickCount()
    If Owner.m_onDiskSigDone Then
        If (now - Owner.m_onDiskCheckedAt) >= 0 And (now - Owner.m_onDiskCheckedAt) < 1000 Then
            BuildOnDiskChanged = Owner.m_onDiskChanged
            Exit Function
        End If
    End If
    Owner.m_onDiskCheckedAt = now
    Dim sig As String
    sig = BuildOnDiskSignature(Owner)
    If Not Owner.m_onDiskSigDone Then
        Owner.m_onDiskSig = sig
        Owner.m_onDiskSigDone = True
        Owner.m_onDiskChanged = False
    Else
        Owner.m_onDiskChanged = (StrComp(sig, Owner.m_onDiskSig, vbBinaryCompare) <> 0)
    End If
    BuildOnDiskChanged = Owner.m_onDiskChanged
End Function

Public Function BuildOnDiskSignature(Owner As CodeTools) As String
    On Error Resume Next
    Dim p3 As String
    p3 = Owner.OwnDllPath()
    If Len(p3) = 0 Then Exit Function
    BuildOnDiskSignature = CStr(FileLen(p3)) & "|" & Format$(FileDateTime(p3), "yyyy-mm-dd hh:nn:ss")
End Function

Public Function BuildOps(Owner As CodeTools, ByVal edits As Object) As String()
    Dim n As Long, ops() As String, i As Long
    n = Owner.CollectionCount(edits)
    ' "no edits" is CALLER INPUT, not an add-in fault. It used to be a bare raise
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
        op = UCase$(McpCls.DictStr(Owner, d, "op"))
        anc = McpCls.DictStr(Owner, d, "anchor")
        txt = Replace(McpCls.DictStr(Owner, d, "text"), vbCrLf, Chr$(1))
        occ = JsonGetLong(d, "occurrence", 0)
        mch = McpCls.DictStr(Owner, d, "match")
        ' gamma-5/6: occurrence (1-based) and match ride along in the op
        ' entry so the resolver can pick the right hit for a repeated line.
        ' every free-text field is escaped so a TAB in it cannot split the
        ' record (occurrence stays the 4th field, so CLng below can never see text).
        ops(i) = op & vbTab & Owner.OpEsc(anc) & vbTab & Owner.OpEsc(txt) & vbTab & CStr(occ) & vbTab & Owner.OpEsc(mch)
        i = i + 1
    Next
    BuildOps = ops
End Function

Public Function BuildPath(Owner As CodeTools) As String
    On Error Resume Next
    BuildPath = Owner.OwnDllPath()
End Function

Public Function BuildSha256(Owner As CodeTools) As String
    If Owner.m_buildShaDone Then
        BuildSha256 = Owner.m_buildSha
        Exit Function
    End If
    Owner.m_buildShaDone = True
    On Error Resume Next
    Dim p As String
    p = Owner.OwnDllPath()
    If Len(p) > 0 Then Owner.m_buildSha = Encoding.Sha256HexFile(p)
    Err.Clear
    BuildSha256 = Owner.m_buildSha
End Function

