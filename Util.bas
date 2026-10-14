Attribute VB_Name = "Util"
Option Explicit

' Small shared helpers: logging, port file, temp dir.

' ---------------------------------------------------------------------------
' Contract v1.3.0 product metadata - the SINGLE source of truth for the version
' stamp. McpServer reports it in initialize (serverInfo.version +
' contractVersion); CodeTools reports it in mcp_info and in the per-response
' mcp stamp. ASCII only: VB6 compiles these sources as ANSI(GBK).
' ---------------------------------------------------------------------------
Public Const VB6MCP_PRODUCT As String = "vb6-mcp"

' TASK-P1-037: this module needs the pid for the instance registry; declare it with
' a module-unique name (CodeTools declares GetCurrentProcessId privately too).
Private Declare Function UtilProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function PumpTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Public Const VB6MCP_PRODUCT_VERSION As String = "0.3.0"

' TASK-P1-042: THE contract banner - a single source of truth. The first line of
' initialize.instructions and mcp_info.contract BOTH read this one constant, so bumping
' the contract is a ONE-LINE change here. NEVER hardcode a contract version anywhere
' else in the product: a stale self-report is the exact defect this card fixed (the
' banner still claimed v1.3.0 after the contract text had reached v1.4.0, and because
' the hub forwards the LIVE banner, users saw the old version).
' TASK-P1-089: the P1-087 snapshot parameter is an inputSchema change, so the contract
' version moves with it. It is ADDITIVE (optional parameter, default off), but a tool-face
' change still has to be named by a version - the rule above is "bump it HERE, in one line".
Public Const VB6MCP_CONTRACT_VERSION As String = "v1.11.0"
' TASK-P1-066: message-pump period and state. The DECLARATIONS live here (VB6 only accepts
' module-level declarations before the first procedure); the pump itself is documented at
' the bottom of this module next to PumpReset / MaybePump / PumpNow.
Public Const PUMP_EVERY_ITER As Long = 1000
Public Const PUMP_EVERY_MS As Long = 150
Private Const PUMP_TICK_MASK As Long = 63

Private m_pumpIter As Long
Private m_pumpTimeIter As Long
Private m_pumpLastAt As Long
Private m_pumpDepth As Long
Private m_pumpCount As Long
Private m_pumpMs As Long
Private m_pumpMaxMs As Long

' TASK-P1-121 item 6 (G-1/G-5): the side-effects AUDIT BLOCK is accumulated per top-level
' call in CodeTools.cls, but VbFmt.bas (format_code / rename_symbol) writes .bak files as
' well and is a different module. This ONE collection is therefore shared: CodeTools
' resets it once per call and reads it when it builds the additive `sideEffects` block;
' VbFmt contributes to it through SeNoteBackup below. Additive only - nothing else in
' this module changed. (The other five collections live in CodeTools.cls because only
' CodeTools creates components / renames them / removes them / writes undo snapshots.)
Public g_seBackups As Collection

Public Sub LogLine(ByVal msg As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp.log"
    Dim ts As Object
    Set ts = fso.OpenTextFile(path, 8, True)
    ts.WriteLine Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & msg
    ts.Close
End Sub

' ---------------------------------------------------------------------------
' Capture the value of AddressOf.
' VB6's grammar rejects "p = AddressOf Foo" (syntax error), and AddressOf may not
' be written into a UDT member either. It IS accepted as an argument to an
' ordinary VB function, so this one-liner is the supported way to obtain the
' address as a Long:
'     Dim p As Long
'     p = FuncAddr(AddressOf MyProc)
' Verified here by compiling AND running a minimal Standard EXE: the captured
' value was 0x4016B0 (a real code address), with the target still callable.
' ---------------------------------------------------------------------------
Public Function FuncAddr(ByVal ptr As Long) As Long
    FuncAddr = ptr
End Function

Public Sub WritePortFile(ByVal port As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp_port.txt"
    Dim ts As Object
    Set ts = fso.CreateTextFile(path, True)
    ts.WriteLine CStr(port)
    ts.Close
End Sub

' ---------------------------------------------------------------------------
' TASK-P1-037 (contract v1.3.1 H2): instance registration
' %TEMP%\vb6mcp\instances\<pid>.json  {pid,port,project,projectPath,dllSha256,
' startedAt,startedByHub}. Written on listen success, deleted on normal exit.
' Stale entries are the READER@s business (pid alive + port listening); a writer
' never deletes another instance@s file.
' ---------------------------------------------------------------------------

Public Function Vb6mcpDir() As String
    Vb6mcpDir = TempDir() & "\vb6mcp"
End Function

Public Function InstancesDir() As String
    InstancesDir = Vb6mcpDir() & "\instances"
End Function

Public Sub WriteInstanceFile(ByVal pid As Long, ByVal port As Long, ByVal project As String, _
                             ByVal projectPath As String, ByVal dllSha As String, _
                             Optional ByVal startedAt As String = "")
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    EnsureFolder fso, Vb6mcpDir()
    EnsureFolder fso, InstancesDir()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("pid") = pid
    d("port") = port
    d("project") = project
    d("projectPath") = projectPath
    d("dllSha256") = dllSha
    ' TASK-P1-064: the entry may be RE-published when the identity changes; startedAt must
    ' keep meaning "when this instance started", so the caller passes the first value back in.
    If Len(startedAt) = 0 Then
        d("startedAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Else
        d("startedAt") = startedAt
    End If
    d("startedByHub") = (Environ$("VB6MCP_STARTED_BY_HUB") = "1")
    Dim ts As Object
    Set ts = fso.CreateTextFile(InstancesDir() & "\" & CStr(pid) & ".json", True)
    ts.WriteLine JsonStringify(d)
    ts.Close
End Sub

Public Sub DeleteInstanceFile(ByVal pid As Long)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    Dim path As String
    path = InstancesDir() & "\" & CStr(pid) & ".json"
    If fso.FileExists(path) Then fso.DeleteFile path, True
End Sub

' Structured bind conflict: a second IDE on the SAME explicit port must fail loudly.
Public Sub WriteBindConflict(ByVal reason As String, ByVal tried As String, ByVal advice As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso Is Nothing Then Exit Sub
    EnsureFolder fso, Vb6mcpDir()
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("pid") = UtilProcessId()
    d("tried") = tried
    d("advice") = advice
    d("at") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Dim ts As Object
    Set ts = fso.CreateTextFile(Vb6mcpDir() & "\bind_conflict.json", True)
    ts.WriteLine JsonStringify(d)
    ts.Close
End Sub

Private Sub EnsureFolder(ByVal fso As Object, ByVal path As String)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If fso.FolderExists(path) Then Exit Sub
    Dim parent As String, i As Long
    i = InStrRev(path, "\")
    If i > 1 Then
        parent = Left$(path, i - 1)
        If Not fso.FolderExists(parent) Then EnsureFolder fso, parent
    End If
    fso.CreateFolder path
End Sub

Public Sub ClearPortFileIfOurs(ByVal port As Long)
    ' delete the port file only when it still points at our port
    ' (another VB6 instance may have written its own port meanwhile)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim path As String
    path = TempDir() & "\vb6mcp_port.txt"
    If fso.FileExists(path) Then
        Dim ts As Object
        Set ts = fso.OpenTextFile(path, 1)
        Dim s As String
        s = Trim$(ts.ReadLine)
        ts.Close
        If s = CStr(port) Then fso.DeleteFile path, True
    End If
End Sub

Public Function TempDir() As String
    TempDir = Environ$("TEMP")
    If TempDir = "" Then TempDir = Environ$("TMP")
    If TempDir = "" Then TempDir = "C:\"
End Function

Public Function FindBarControl(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String) As Object
    ' Two passes, because the menu bar NAME is localized: on a Chinese IDE it
    ' is the local word for "menu bar", so a bar-name filter never matches and
    ' the whole lookup died before reaching any control - which surfaced as
    ' "Menu command not found: Start|..." even when the item existed.
    '   1) restrict to bars whose name matches barCands (precise, fast)
    '   2) ignore the bar name entirely
    Set FindBarControl = FindBarControlPass(vbe, barCands, ctrlCands, True)
    If Not FindBarControl Is Nothing Then Exit Function
    Set FindBarControl = FindBarControlPass(vbe, barCands, ctrlCands, False)
End Function

Public Function FindBarControlPass(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String, ByVal filterBar As Boolean) As Object
    ' Returns the matching CommandBarControl (searching bars and one level of
    ' popup sub-menus) or Nothing.
    ' Candidates are pipe-separated names covering English and localized
    ' (Chinese) captions; "(&x)" accelerators are ignored and candidates
    ' also prefix-match (e.g. "Make" matches "Make Project1.exe...").
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long
    nb = bars.Count
    Dim bi As Long
    Dim b As Object
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            If (Not filterBar) Or MatchCand(b.Name, barCands) Then
                Dim nc As Long
                nc = b.Controls.Count
                Dim ci As Long
                Dim c As Object
                For ci = 1 To nc
                    Set c = b.Controls.Item(ci)
                    If Not c Is Nothing Then
                        If MatchCand(c.Caption, ctrlCands) Then
                            Set FindBarControlPass = c
                            Exit Function
                        End If
                        ' one level of popup sub-menus (buttons have no
                        ' child controls: touching .Controls on them
                        ' raises 438 - guarded access)
                        Dim subc As Object
                        Dim subCol As Object
                        Err.Clear
                        Set subCol = c.Controls
                        If Not subCol Is Nothing Then
                            Dim ns As Long
                            ns = subCol.Count
                            Dim si As Long
                            For si = 1 To ns
                                Set subc = subCol.Item(si)
                                If Not subc Is Nothing Then
                                    If MatchCand(subc.Caption, ctrlCands) Then
                                        Set FindBarControlPass = subc
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Function

Public Function MatchCand(ByVal caption As String, ByVal cands As String) As Boolean
    Dim base As String
    base = caption
    base = Replace(base, "&", "")
    Dim p As Long
    p = InStr(base, "(")
    If p > 0 Then base = Trim$(Left$(base, p - 1))
    If Len(base) = 0 Then Exit Function
    Dim pos As Long
    Dim sep As Long
    Dim cand As String
    pos = 1
    Do While pos <= Len(cands)
        sep = InStr(pos, cands, "|")
        If sep = 0 Then sep = Len(cands) + 1
        cand = Trim$(Mid$(cands, pos, sep - pos))
        If Len(cand) > 0 Then
            cand = Replace(cand, "&", "")
            If StrComp(base, cand, vbTextCompare) = 0 Then
                MatchCand = True
                Exit Function
            End If
            If Len(base) >= Len(cand) Then
                If StrComp(Left$(base, Len(cand)), cand, vbTextCompare) = 0 Then
                    MatchCand = True
                    Exit Function
                End If
            End If
        End If
        pos = sep + 1
    Loop
End Function

' ---------------------------------------------------------------------------
' Card TASK-P1-007 (i18n): Id-first command lookup.
'
' Until now a menu item could only be identified by its LOCALIZED caption, so a
' VB6 IDE in any language other than English/Chinese lost every menu command.
' The CommandBars model also exposes a language-independent Id - but TASK-P1-006
' measured that VB6 Ids are NOT unique: id=746 alone covers 986 nodes (including
' both "Make" and "Start With Full Compile"), and Start/Continue share id=186.
' A bare Id query would therefore hit the wrong control.
'
' Rule enforced here (per S's ruling on REQ-20260921-016): the Id is NEVER used
' on its own. It is AND-ed with the shortcut text when one is known, and
' otherwise with the caption candidates. If no Id-verified candidate is found we
' fall back to the original localized-caption lookup, so the pre-existing
' candidate lists keep working and nothing that used to be found is lost.
' ---------------------------------------------------------------------------
Public Function FindBarControlById(ByVal vbe As Object, ByVal barCands As String, _
                                   ByVal ctrlCands As String, ByVal wantId As Long, _
                                   ByVal wantShortcut As String) As Object
    ' Id-first, using the SAME two passes as the caption matcher: first scoped to
    ' bars matching barCands (cheap), then across every bar. Scope matters - an
    ' unscoped walk has to touch all ~2300 CommandBars nodes on every call.
    If wantId <> 0 Then
        Dim v As Object
        Set v = FindByIdPass(vbe, barCands, ctrlCands, wantId, wantShortcut, True)
        If Not v Is Nothing Then
            Set FindBarControlById = v
            Exit Function
        End If
        Set v = FindByIdPass(vbe, barCands, ctrlCands, wantId, wantShortcut, False)
        If Not v Is Nothing Then
            Set FindBarControlById = v
            Exit Function
        End If
    End If
    ' last fallback: the localized-caption matcher, unchanged
    Set FindBarControlById = FindBarControl(vbe, barCands, ctrlCands)
End Function

Private Function FindByIdPass(ByVal vbe As Object, ByVal barCands As String, ByVal ctrlCands As String, _
                              ByVal wantId As Long, ByVal wantShortcut As String, _
                              ByVal filterBar As Boolean) As Object
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long
    Dim bi As Long
    Dim b As Object
    nb = bars.Count
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            If (Not filterBar) Or MatchCand(b.Name, barCands) Then
                Dim nc As Long
                Dim ci As Long
                Dim c As Object
                nc = b.Controls.Count
                For ci = 1 To nc
                    Set c = b.Controls.Item(ci)
                    If Not c Is Nothing Then
                        If CtrlMatches(c, wantId, wantShortcut, ctrlCands) Then
                            Set FindByIdPass = c
                            Exit Function
                        End If
                        ' one level of popup sub-menus; buttons raise 438 on .Controls
                        Dim subCol As Object
                        Set subCol = Nothing
                        Err.Clear
                        Set subCol = c.Controls
                        If Not subCol Is Nothing Then
                            Dim ns As Long
                            Dim si As Long
                            Dim subc As Object
                            ns = subCol.Count
                            For si = 1 To ns
                                Set subc = subCol.Item(si)
                                If Not subc Is Nothing Then
                                    If CtrlMatches(subc, wantId, wantShortcut, ctrlCands) Then
                                        Set FindByIdPass = subc
                                        Exit Function
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
End Function

Private Function CtrlMatches(ByVal c As Object, ByVal wantId As Long, _
                             ByVal wantShortcut As String, ByVal ctrlCands As String) As Boolean
    On Error Resume Next
    Dim gId As Long
    gId = -1
    gId = c.Id
    If gId <> wantId Then Exit Function
    If Len(wantShortcut) > 0 Then
        Dim sh As String
        sh = ""
        sh = c.ShortcutText
        If StrComp(Trim$(sh), Trim$(wantShortcut), vbTextCompare) <> 0 Then Exit Function
    Else
        ' No shortcut is known for this command, so the caption must still agree.
        ' This is what makes an ambiguous Id such as 746 safe to refuse.
        If Not MatchCand(c.Caption, ctrlCands) Then Exit Function
    End If
    CtrlMatches = True
End Function

' ---------------------------------------------------------------------------
' TASK-P1-069: Id-ONLY command lookup, used by command_bars when the caller
' addresses an item as "#746" / "id:746" / "746".
'
' The existing FindBarControlById is Id-AND-(shortcut or caption), because a
' bare Id is ambiguous. A command_bars caller who only knows the Id has no
' caption to AND with, so this variant answers the ambiguity openly: it returns
' a control ONLY when exactly ONE ENABLED control matches, and otherwise leaves
' matchCount and a short candidate list for the caller to report. Refusing to
' guess is the feature, not a limitation.
' ---------------------------------------------------------------------------
Public Function FindBarControlByIdOnly(ByVal vbe As Object, ByVal barCands As String, _
        ByVal capCands As String, ByVal wantId As Long, _
        ByRef matchCount As Long, ByRef candsOut As String) As Object
    On Error Resume Next
    matchCount = 0
    candsOut = ""
    Dim v As Object
    If Len(barCands) > 0 Then
        ' pass 1: the caller's bar argument must match the BAR NAME or a node on
        ' the control's menu PATH (e.g. bar="工程" for 工程>部件).
        Set v = FindByIdOnlyPass(vbe, barCands, capCands, wantId, matchCount, candsOut)
        If Not v Is Nothing Then
            Set FindBarControlByIdOnly = v
            Exit Function
        End If
        If matchCount > 0 Then Exit Function   ' ambiguous inside that scope: do not widen
        ' pass 2: the bar name is localized, so retry with no bar scope
        Set v = FindByIdOnlyPass(vbe, "", capCands, wantId, matchCount, candsOut)
    Else
        Set v = FindByIdOnlyPass(vbe, "", capCands, wantId, matchCount, candsOut)
    End If
    If Not v Is Nothing Then Set FindBarControlByIdOnly = v
End Function

Private Function FindByIdOnlyPass(ByVal vbe As Object, ByVal barCands As String, ByVal capCands As String, _
        ByVal wantId As Long, ByRef matchCount As Long, ByRef candsOut As String) As Object
    On Error Resume Next
    Dim bars As Object
    Set bars = vbe.CommandBars
    If bars Is Nothing Then Exit Function
    Dim nb As Long, bi As Long, b As Object, hit As Object, under As Boolean
    nb = bars.Count
    Set hit = Nothing
    For bi = 1 To nb
        Set b = bars.Item(bi)
        If Not b Is Nothing Then
            under = True
            If Len(barCands) > 0 Then under = MatchCand(b.Name, barCands)
            FindByIdOnlyIn b.Controls, b.Name, 0, wantId, barCands, capCands, under, hit, matchCount, candsOut
        End If
    Next
    If matchCount = 1 Then Set FindByIdOnlyPass = hit
End Function

' Qualify = Id matches AND enabled AND (the caller's bar argument matched the bar
' name or an ancestor on the menu PATH) AND (when the caller also gave a caption,
' the caption matches). The caption conjunct is REQUIRED: TASK-P1-069 measured
' that id=746 has 525 ENABLED matches across the menus and still 6 inside the
' 工程 submenu alone, so an id on its own cannot name a control.
Private Sub FindByIdOnlyIn(ByVal col As Object, ByVal barName As String, ByVal depth As Long, _
        ByVal wantId As Long, ByVal barCands As String, ByVal capCands As String, ByVal under As Boolean, _
        ByRef hit As Object, ByRef matchCount As Long, ByRef candsOut As String)
    On Error Resume Next
    If col Is Nothing Then Exit Sub
    If depth > 2 Then Exit Sub
    Dim n As Long, k As Long, c As Object, subCol As Object
    Dim selfOk As Boolean, cap As String
    n = col.Count
    For k = 1 To n
        Set c = col.Item(k)
        If Not c Is Nothing Then
            cap = ""
            cap = c.Caption
            selfOk = under
            If Len(barCands) = 0 Then selfOk = True
            If Not selfOk Then selfOk = MatchCand(cap, barCands)
            Dim gid As Long, en As Boolean, qual As Boolean
            gid = -1
            en = False
            gid = c.Id
            en = c.Enabled
            qual = (gid = wantId) And en And selfOk
            If qual And Len(capCands) > 0 Then qual = MatchCand(cap, capCands)
            If qual Then
                matchCount = matchCount + 1
                If hit Is Nothing Then Set hit = c
                If Len(candsOut) < 300 Then
                    If Len(candsOut) > 0 Then candsOut = candsOut & "; "
                    candsOut = candsOut & barName & ">" & cap
                End If
            End If
            Set subCol = Nothing
            Err.Clear
            Set subCol = c.Controls
            If Not subCol Is Nothing Then
                FindByIdOnlyIn subCol, barName, depth + 1, wantId, barCands, capCands, selfOk, hit, matchCount, candsOut
            End If
        End If
    Next
End Sub


' (kernel32 GetTickCount is declared in the declarations section above - VB6
'  only accepts Declare statements before the first procedure.)
' ===========================================================================
' TASK-P1-066: THE unified message-pump point.
'
' WHY. VB6 hosts this add-in on the IDE's single UI thread (STA). While a tool
' call computes, the IDE message queue is not serviced, so the IDE UI - and
' every other MCP request - is frozen until the call returns. Measured BEFORE
' this fix on a 3 x 26,999-line fixture: a 3.2 s verify_scope made a concurrent
' ide_state wait 2.9 s and return at the very instant the long call ended
' (内部证据). Long loops therefore hand the
' queue back periodically.
'
' WHAT. Exactly ONE helper. Rule enforced by grep: no module may call DoEvents
' directly (the only other site is the pre-existing control_execution wait,
' which is deliberate and documented there). Every pump point is the literal
' call MaybePump, so the set of pump points is auditable in one grep.
'
' PERIOD. Every PUMP_EVERY_ITER iterations OR every PUMP_EVERY_MS milliseconds,
' whichever comes first - the user asked for a PERIOD, not a call per tick.
' Both are named constants so the policy can be retuned in one place. The tick
' is only sampled every 64 iterations, so the per-iteration cost is two Long
' increments.
'
' RE-ENTRANCY. A pump never nests (m_pumpDepth): once DoEvents is running, a
' second DoEvents would recurse the message loop. The tool-level guard lives in
' CodeTools.CallTool, which refuses a second state-changing call while one is
' in flight.
'
' NOT FOR. Error handlers, and critical sections covered by On Error Resume
' Next whose state must not be disturbed.
' ===========================================================================


' Start a fresh period window (called once per top-level tool call).
Public Sub PumpReset()
    m_pumpIter = 0
    m_pumpTimeIter = 0
    m_pumpLastAt = PumpTickCount()
End Sub

Public Sub MaybePump()
    ' A pump inside a pump must never happen.
    If m_pumpDepth > 0 Then Exit Sub
    m_pumpIter = m_pumpIter + 1
    If m_pumpIter >= PUMP_EVERY_ITER Then
        PumpNow
        Exit Sub
    End If
    m_pumpTimeIter = m_pumpTimeIter + 1
    If (m_pumpTimeIter And PUMP_TICK_MASK) = 0 Then
        If (PumpTickCount() - m_pumpLastAt) >= PUMP_EVERY_MS Then PumpNow
    End If
End Sub

Public Sub PumpNow()
    m_pumpIter = 0
    m_pumpTimeIter = 0
    m_pumpDepth = m_pumpDepth + 1
    m_pumpCount = m_pumpCount + 1
    Dim t0 As Long
    t0 = PumpTickCount()
    m_pumpLastAt = t0
    DoEvents
    ' TASK-P1-067: throttle-rescan the optional code-pane wheel. It returns
    ' immediately while the wheel is OFF (the default), so the pump cost is
    ' unchanged by default and no hook is ever installed.
    WinOps.WheelTick
    Dim dt As Long
    dt = PumpTickCount() - t0
    m_pumpMs = m_pumpMs + dt
    If dt > m_pumpMaxMs Then m_pumpMaxMs = dt
    m_pumpDepth = m_pumpDepth - 1
End Sub

Public Function PumpCount() As Long
    PumpCount = m_pumpCount
End Function

Public Function PumpMsTotal() As Long
    PumpMsTotal = m_pumpMs
End Function

Public Function PumpMsMax() As Long
    PumpMsMax = m_pumpMaxMs
End Function

Public Function PumpStatsJson() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("calls") = m_pumpCount
    d("msInPump") = m_pumpMs
    d("maxPumpMs") = m_pumpMaxMs
    d("iterPeriod") = PUMP_EVERY_ITER
    d("msPeriod") = PUMP_EVERY_MS
    PumpStatsJson = JsonStringify(d)
End Function

' Read-and-reset, for a measurement run.
Public Function PumpStatsConsume() As String
    Dim s As String
    s = PumpStatsJson()
    m_pumpCount = 0
    m_pumpMs = 0
    m_pumpMaxMs = 0
    PumpStatsConsume = s
End Function

' TASK-P1-121 item 6: record a .bak this call created (VbFmt.bas calls it). The collection
' itself is declared in the declarations section above; CodeTools resets it per call.
Public Sub SeNoteBackup(ByVal path As String)
    On Error Resume Next
    If Len(path) = 0 Then Exit Sub
    If g_seBackups Is Nothing Then Set g_seBackups = New Collection
    g_seBackups.Add path
    Err.Clear
End Sub
