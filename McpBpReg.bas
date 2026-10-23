Attribute VB_Name = "McpBpReg"
Option Explicit

' =======================================================================
' C-3 facade batch: McpBpReg (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Private Const BP_REG_DIR As String = "vb6mcp"
Private Const BP_REG_FILE As String = "breakpoints.json"

Public Function BpProbeLedgerCount(Owner As CodeTools) As Long
    ' ONE count for the probe cleanup proof: the disk ledger plus this session's records.
    On Error Resume Next
    Dim coll As Object
    Dim n As Long
    n = 0
    Set coll = BpRegLoad(Owner)
    If Not coll Is Nothing Then n = coll.Count
    If Not Owner.m_bpSession Is Nothing Then n = n + Owner.m_bpSession.Count
    BpProbeLedgerCount = n
    Err.Clear
End Function

Public Function BpPruneJson(Owner As CodeTools, ByRef root As Object) As String
    Dim disk As Object
    Set disk = BpRegLoad(Owner)
    Dim keep As Object
    Set keep = New Collection
    Dim removed As Object
    Set removed = New Collection
    Dim i As Long
    For i = 1 To disk.Count
        Dim e As Object
        Set e = disk(i)
        Dim comp As String
        Dim ln As Long
        comp = JsonGetStr(e, "component", "")
        ln = JsonGetLong(e, "paneLine", 0)
        Dim cls As Object
        Set cls = McpBpPane.BpClassifyPaneLine(Owner, comp, ln)
        Dim refused As Boolean
        refused = (Not CBool(cls("breakable")))
        Dim shaNow As String
        shaNow = McpBpPane.BpPaneSha(Owner, comp)
        Dim isStale As Boolean
        isStale = (Len(JsonGetStr(e, "paneShaAtSet", "")) > 0 And StrComp(JsonGetStr(e, "paneShaAtSet", ""), shaNow, vbTextCompare) <> 0)
        If refused Or isStale Then
            removed.Add e
        Else
            keep.Add e
        End If
    Next
    Dim wok As Boolean
    wok = BpRegSaveCollection(Owner, keep)
    root("ok") = wok
    root("ledgerWrite") = Owner.m_bpLedgerStatus
    If Not wok Then
        root("reason") = McpBpCore.BpLedgerReason(Owner)
        root("ledgerWriteError") = McpBpCore.BpLedgerErrText(Owner, Owner.m_bpLedgerStatusDetail)
    End If
    root("prune") = True
    root("removedCount") = removed.Count
    Set root("removed") = removed
    root("keptCount") = keep.Count
    root("note") = "prune only edits THIS tool's disk ledger (breakpoints.json under %TEMP%\vb6mcp by default, or VB6MCP_BP_REG_FILE when set; the legacy %TEMP%\.vb6mcp directory is no longer read or written). Native breakpoints already set inside the IDE cannot be cleared individually (no VBIDE API) - use breakpoints{action:clear_all} if you need all of them gone"
    BpPruneJson = JsonStringify(root)
End Function

Public Function BpRegLoad(Owner As CodeTools) As Object
    Dim out As Object
    Set out = New Collection
    Set BpRegLoad = out
    On Error Resume Next
    Dim p As String
    p = BpRegPath(Owner)
    If Len(p) = 0 Then Exit Function
    Dim ok As Boolean
    Dim txt As String
    txt = Owner.ReadTextUtf8File(p, ok)
    If Not ok Then Exit Function
    If Len(txt) = 0 Then Exit Function
    Dim v As Variant
    If Not JsonParse(txt, v) Then Exit Function
    Dim d As Object
    Set d = v
    If d Is Nothing Then Exit Function
    If Not d.Exists("breakpoints") Then Exit Function
    Dim arr As Object
    Set arr = Nothing
    Set arr = d("breakpoints")
    If arr Is Nothing Then Exit Function
    Dim e As Variant
    For Each e In arr
        If IsObject(e) Then out.Add e
    Next
    Err.Clear
    Set BpRegLoad = out
End Function

Public Function BpRegPath(Owner As CodeTools) As String
    ' 33b S3: the ledger is TOOL STATE, so it must NOT be written into the user's project /
    ' current directory. Default = %TEMP%\vb6mcp\breakpoints.json; VB6MCP_BP_REG_FILE overrides
    ' the whole file path (used by tests and by users who want a custom location).
    BpRegPath = ""
    Dim ov As String
    ov = Trim$(Environ$("VB6MCP_BP_REG_FILE"))
    If Len(ov) > 0 Then
        BpRegPath = ov
        Exit Function
    End If
    Dim tdir As String
    tdir = Environ$("TEMP")
    If Len(tdir) = 0 Then tdir = Environ$("TMP")
    If Len(tdir) = 0 Then Exit Function
    BpRegPath = Owner.PathJoin(Owner.PathJoin(tdir, BP_REG_DIR), BP_REG_FILE)
End Function

' Folder that must exist before the ledger is written (parent of BpRegPath).
Public Function BpRegDirOf(Owner As CodeTools) As String
    BpRegDirOf = ""
    Dim p As String
    p = BpRegPath(Owner)
    If Len(p) = 0 Then Exit Function
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    BpRegDirOf = fso.GetParentFolderName(p)
    Err.Clear
End Function

' Create the ledger folder (and its parent when needed) WITHOUT touching the project dir.
Private Function BpRegEnsureDir(ByVal dd As String) As Boolean
    BpRegEnsureDir = False
    If Len(dd) = 0 Then Exit Function
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(dd) Then
        BpRegEnsureDir = True
        Exit Function
    End If
    Dim par As String
    par = fso.GetParentFolderName(dd)
    If Len(par) > 0 Then
        If Not fso.FolderExists(par) Then fso.CreateFolder par
    End If
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    fso.CreateFolder dd
    If Err.Number = 0 Then BpRegEnsureDir = True
    Err.Clear
End Function

Public Function BpRegSave(Owner As CodeTools) As Boolean
    BpRegSave = False
    Owner.m_bpLedgerStatus = "failed"
    Set Owner.m_bpLedgerStatusDetail = Nothing
    On Error Resume Next
    Dim p As String
    p = BpRegPath(Owner)
    If Len(p) = 0 Then
        Set Owner.m_bpLedgerStatusDetail = McpBpCore.BpLedgerStatusStub(Owner, "ledger_path_unavailable")
        Err.Clear
        Exit Function
    End If
    Dim reg As Object
    Set reg = BpRegistry(Owner)
    Dim arr As Object
    Set arr = New Collection
    Dim k As Variant
    For Each k In reg.Keys
        Dim rec As Object
        Set rec = reg(k)
        Dim e As Object
        Set e = CreateObject("Scripting.Dictionary")
        e("id") = rec("id")
        e("component") = rec("component")
        e("paneLine") = rec("paneLine")
        e("diskLine") = rec("diskLine")
        e("mode") = rec("mode")
        e("enabled") = rec("enabled")
        e("paneShaAtSet") = CStr(rec("paneShaAtSet"))
        e("setAt") = CStr(rec("setAt"))
        If rec.Exists("stale") Then e("stale") = CBool(rec("stale"))
        If rec.Exists("verification") Then e("verification") = CStr(rec("verification"))
        If rec.Exists("staleCheckedAt") Then e("staleCheckedAt") = CStr(rec("staleCheckedAt"))
        arr.Add e
    Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("schema") = 1
    root("tool") = "breakpoints"
    root("savedAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    root("diskBacked") = False
    root("diskBackedMeasured") = False
    root("humanNative") = "NOT_ENUMERABLE_no_VBIDE_API"
    Set root("breakpoints") = arr
    Dim dd As String
    dd = BpRegDirOf(Owner)
    If Not BpRegEnsureDir(dd) Then
        Set Owner.m_bpLedgerStatusDetail = McpBpCore.BpLedgerStatusStub(Owner, "ledger_dir_unavailable")
        Err.Clear
        Exit Function
    End If
    ' 33b S3: the write is a side effect of THIS tool - it must appear in sideEffects.filesWritten.
    Owner.SeNoteFile p
    ' FIX-13: atomic replace + read-back validation. The old code wrote in place and did not
    ' truncate, so a shorter ledger kept the tail of the previous one (byte-level evidence:
    ' 573-byte JSON in a 581-byte file, 8 stale bytes).
    Dim wr As Object
    Set wr = McpBpCore.BpLedgerWriteAtomic(Owner, p, JsonStringify(root), arr.Count)
    Owner.m_bpLedgerStatus = McpBpCore.BpLedgerStatusOf(Owner, wr)
    Owner.m_bpLedgerFailed = (Owner.m_bpLedgerStatus <> "ok")
    If Owner.m_bpLedgerFailed Then
        Set Owner.m_bpLedgerFailedDetail = wr
        Set Owner.m_bpLedgerStatusDetail = wr
    Else
        Set Owner.m_bpLedgerFailedDetail = Nothing
        Set Owner.m_bpLedgerStatusDetail = Nothing
    End If
    If Not Owner.m_bpLedgerFailed Then BpRegSave = True
    Err.Clear
End Function

Public Function BpRegSaveCollection(Owner As CodeTools, ByVal coll As Object) As Boolean
    BpRegSaveCollection = False
    Owner.m_bpLedgerStatus = "failed"
    Set Owner.m_bpLedgerStatusDetail = Nothing
    On Error Resume Next
    Dim p As String
    p = BpRegPath(Owner)
    If Len(p) = 0 Then
        Set Owner.m_bpLedgerStatusDetail = McpBpCore.BpLedgerStatusStub(Owner, "ledger_path_unavailable")
        Err.Clear
        Exit Function
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("schema") = 1
    root("tool") = "breakpoints"
    root("savedAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    root("diskBacked") = False
    root("humanNative") = "NOT_ENUMERABLE_no_VBIDE_API"
    Set root("breakpoints") = coll
    Dim dd As String
    dd = BpRegDirOf(Owner)
    If Not BpRegEnsureDir(dd) Then
        Set Owner.m_bpLedgerStatusDetail = McpBpCore.BpLedgerStatusStub(Owner, "ledger_dir_unavailable")
        Err.Clear
        Exit Function
    End If
    ' 33b S3: the write is a side effect of THIS tool - it must appear in sideEffects.filesWritten.
    Owner.SeNoteFile p
    ' FIX-22i: ALWAYS go through the atomic primitive (temp -> read-back parse/bytes/count check
    ' -> replace). The old direct WriteTextUtf8File opened For Binary without truncation, so a
    ' shorter ledger kept the tail of the previous, longer one (illegal JSON on disk).
    Dim wr As Object
    Set wr = McpBpCore.BpLedgerWriteAtomic(Owner, p, JsonStringify(root), coll.Count)
    Owner.m_bpLedgerStatus = McpBpCore.BpLedgerStatusOf(Owner, wr)
    Owner.m_bpLedgerFailed = (Owner.m_bpLedgerStatus <> "ok")
    If Owner.m_bpLedgerFailed Then
        Set Owner.m_bpLedgerFailedDetail = wr
        Set Owner.m_bpLedgerStatusDetail = wr
    Else
        Set Owner.m_bpLedgerFailedDetail = Nothing
        Set Owner.m_bpLedgerStatusDetail = Nothing
    End If
    BpRegSaveCollection = (Owner.m_bpLedgerStatus = "ok")
    Err.Clear
End Function

Public Function BpRegistry(Owner As CodeTools) As Object
    If Owner.m_bpSession Is Nothing Then Set Owner.m_bpSession = CreateObject("Scripting.Dictionary")
    Set BpRegistry = Owner.m_bpSession
End Function

Public Sub BpRemoveLedgerId(Owner As CodeTools, ByVal id As String)
    On Error Resume Next
    Dim reg As Object
    Set reg = BpRegistry(Owner)
    If reg.Exists(id) Then reg.Remove id
    Dim disk As Object
    Set disk = BpRegLoad(Owner)
    Dim keep As Object
    Set keep = New Collection
    Dim seen As Object
    Set seen = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 1 To disk.Count
        Dim e As Object
        Set e = disk(i)
        Dim did As String
        did = JsonGetStr(e, "id", "")
        If Len(did) > 0 Then
            If StrComp(did, id, vbTextCompare) <> 0 Then
                keep.Add e
                seen(did) = True
            End If
        End If
    Next
    Dim k As Variant
    For Each k In reg.Keys
        Dim rec As Object
        Set rec = reg(k)
        Dim rid As String
        rid = CStr(rec("id"))
        If Len(rid) > 0 Then
            If Not seen.Exists(rid) Then
                keep.Add rec
                seen(rid) = True
            End If
        End If
    Next
    Call BpRegSaveCollection(Owner, keep)
    Err.Clear
End Sub

