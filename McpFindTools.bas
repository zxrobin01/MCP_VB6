Attribute VB_Name = "McpFindTools"
Option Explicit

' =======================================================================
' C-3 facade batch: McpFindTools (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Public Function FindAnchor(Owner As CodeTools, ByRef a() As String, ByRef b() As String, ByVal alo As Long, ByVal ahi As Long, ByVal blo As Long, ByVal bhi As Long, ByRef ai As Long, ByRef bi As Long) As Boolean
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

Public Function FindBarControlDeep(Owner As CodeTools, ByVal ctrlCands As String) As Object
    On Error Resume Next
    Dim bars As Object, b As Object, i As Long
    Set bars = Owner.m_vbe.CommandBars
    If bars Is Nothing Then Exit Function
    For i = 1 To bars.Count
        Set b = bars.Item(i)
        If Not b Is Nothing Then
            Set FindBarControlDeep = FindControlDeep(Owner, b.Controls, ctrlCands, 0)
            If Not FindBarControlDeep Is Nothing Then Exit Function
        End If
    Next
End Function

Public Function FindBusRoot(Owner As CodeTools) As String
    FindBusRoot = ""
    Dim d As String, i As Long, p As Long
    d = McpProjTools.ProjectDir(Owner)
    If Len(d) = 0 Then d = App.Path
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    If fso Is Nothing Then Exit Function
    For i = 0 To 6
        If Len(d) = 0 Then Exit For
        On Error Resume Next
        If fso.FolderExists(d & "\\tasks") Then
            FindBusRoot = d
            Exit Function
        End If
        On Error GoTo 0
        p = InStrRev(d, "\")
        If p <= 3 Then Exit For
        d = Left$(d, p - 1)
    Next
End Function

Public Function FindCmd(Owner As CodeTools, ByVal key As String, ByVal barCands As String, ByVal ctrlCands As String) As Object
    Dim sc As String
    Dim i As Long
    sc = ""
    i = Owner.CmdId(key, sc)
    Set FindCmd = FindBarControlById(Owner.m_vbe, barCands, ctrlCands, i, sc)
End Function

Public Function FindCodePaneByCaption(Owner As CodeTools, ByVal caption As String) As VBIDE.CodePane
    On Error Resume Next
    Set FindCodePaneByCaption = Nothing
    If Len(caption) = 0 Then Exit Function
    Owner.m_wheelResolveVia = "none"
    Dim cp As VBIDE.CodePane
    ' (1) verbatim Window.Caption match - the DESIGN-time channel, unchanged.
    For Each cp In Owner.m_vbe.CodePanes
        If Not cp Is Nothing Then
            If Not cp.Window Is Nothing Then
                If StrComp(cp.Window.Caption, caption, vbBinaryCompare) = 0 Then
                    Set FindCodePaneByCaption = cp
                    Owner.m_wheelResolveVia = "caption"
                    Exit Function
                End If
            End If
        End If
    Next
    ' (2) (measured 2026-10-13: in BREAK mode CodePane.Window is
    ' not readable at all - get_codepane{active} answers TopLine/CodeModule but its
    ' windowCaption comes back EMPTY - so channel (1) can never match. Identify the pane
    ' by the COMPONENT NAME embedded in the window title ("<project> - <name> (Code)")
    ' and reach it through the SAME object path the active-pane tool uses
    ' (CodeModule.CodePane), which is readable in break.
    Dim nm As String
    nm = Owner.ComponentNameFromPaneCaption(caption)
    If Len(nm) > 0 Then
        Dim c2 As VBIDE.VBComponent
        Set c2 = FindComponent(Owner, nm)
        If Not c2 Is Nothing Then
            If Not c2.CodeModule Is Nothing Then
                Set FindCodePaneByCaption = c2.CodeModule.CodePane
                If Not FindCodePaneByCaption Is Nothing Then
                    Owner.m_wheelResolveVia = "component"
                    Exit Function
                End If
            End If
        End If
    End If
    ' (3) last resort: the ACTIVE pane, but ONLY when the caption really names it, so a
    ' non-code window (Immediate / Watch / Locals) can never be mapped onto it - that
    ' mapping is exactly the residual this task forbids.
    Dim ap As VBIDE.CodePane
    Set ap = Owner.m_vbe.ActiveCodePane
    If Not ap Is Nothing Then
        Dim an As String
        an = ""
        If Not ap.CodeModule Is Nothing Then an = ap.CodeModule.Name
        If Len(an) > 0 Then
            If InStr(1, caption, an, vbTextCompare) > 0 Then
                Set FindCodePaneByCaption = ap
                Owner.m_wheelResolveVia = "active_name"
            End If
        End If
    End If
    Err.Clear
End Function

Public Function FindComponent(Owner As CodeTools, ByVal name As String) As VBIDE.VBComponent
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    Dim comp As VBIDE.VBComponent
    For Each comp In p.VBComponents
        If LCase$(comp.Name) = LCase$(name) Then
            Set FindComponent = comp
            Exit Function
        End If
    Next
    ' "the component you named is not in this project" is CALLER INPUT, not an
    ' add-in fault. It used to be a bare raise (+20) that CallToolInner could only wrap as
    ' ok:false / reason:"unexpected_error" / vbErrorCode:-2147221484 for EVERY one of the 36
    ' call sites (41 tools once the helpers are counted). Capture the diagnostics the caller
    ' needs, then tag the raise so the MCPREASON channel answers with a branchable refusal.
    ' This is the ONLY place affected: the success path (component found) does not change at
    ' all, and no other refusal is touched.
    Owner.CnfCapture name, p
    ' FIX-24 (F1): a Reload RENAMED the component that used to carry this name moments ago, so this
    ' miss is the TRANSIENT reload window, not a wrong name. Answer pane_not_settled (+ the name it
    ' resolved to) so a caller retries instead of fixing a name that was right.
    Dim settleTo As String
    Dim settleAge As Long
    If Owner.ReloadRenameJustHappened(name, settleTo, settleAge) Then
        Owner.m_settleFrom = name
        Owner.m_settleTo = settleTo
        Owner.m_settleAgeMs = settleAge
        Err.Raise vbObjectError + 21, "CodeTools", "MCPREASON:pane_not_settled"
    End If
    Err.Raise vbObjectError + 20, "CodeTools", "MCPREASON:component_not_found"
End Function

Public Function FindComponentExactType(Owner As CodeTools, ByVal nm As String, ByVal ct As vbext_ComponentType) As VBIDE.VBComponent
    Set FindComponentExactType = Nothing
    If Len(nm) = 0 Then Exit Function
    On Error Resume Next
    Dim i As Long
    Dim c As VBIDE.VBComponent
    For i = 1 To Owner.GetProject().VBComponents.Count
        Set c = Nothing
        Set c = Owner.GetProject().VBComponents.Item(i)
        If Not c Is Nothing Then
            If StrComp(c.Name, nm, vbTextCompare) = 0 Then
                If c.Type = ct Then Set FindComponentExactType = c
            End If
        End If
    Next
    Err.Clear
End Function

Public Function FindControlDeep(Owner As CodeTools, ByVal col As Object, ByVal ctrlCands As String, ByVal depth As Long) As Object
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
                Set FindControlDeep = FindControlDeep(Owner, subCol, ctrlCands, depth + 1)
                If Not FindControlDeep Is Nothing Then Exit Function
            End If
        End If
    Next
End Function

Public Sub FindHitsIdx(Owner As CodeTools, ByVal text As String, ByRef iRaw() As String, ByRef iNorm() As String, ByRef iFirst As Object, ByRef iCnts As Object, ByRef iRawFirst As Object, ByRef iRawCnts As Object, ByVal anchor As String, ByVal matchMode As String, ByVal scopeStart As Long, ByVal scopeEnd As Long, ByRef hits() As Long, ByRef cnt As Long)
    Dim exact As Boolean, i As Long, lo As Long, hi As Long, k As String, nk As String
    exact = (LCase$(Trim$(matchMode)) = "exact")
    If exact Then Owner.EnsureRawIndex iRaw, iRawFirst, iRawCnts
    ' / FindHitsEx: an EMPTY pane has NO hits (Split("") is an empty array).
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
        hi = McpCls.LineUbound(Owner, iRaw) + 1
        If scopeEnd > 0 And scopeEnd < hi Then hi = scopeEnd
        ReDim hits(0 To McpCls.LineUbound(Owner, iRaw))
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
    ReDim hits(0 To McpCls.LineUbound(Owner, iRaw))
    cnt = 0
    If exact Then
        For i = 0 To McpCls.LineUbound(Owner, iRaw)
            If StrComp(iRaw(i), anchor, vbBinaryCompare) = 0 Then
                hits(cnt) = i + 1
                cnt = cnt + 1
            End If
        Next
    Else
        For i = 0 To McpCls.LineUbound(Owner, iRaw)
            If iNorm(i) = k Then
                hits(cnt) = i + 1
                cnt = cnt + 1
            End If
        Next
    End If
End Sub

Public Function FindMenuBlockLines(Owner As CodeTools, ByRef lines() As String, ByVal menuName As String, ByRef b1 As Long, ByRef b2 As Long, ByRef ind As Long) As Boolean
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

Public Function FindNoHitsJson(Owner As CodeTools, ByVal component As String, ByVal requested As Object, ByVal countOfLines As Long, ByVal defaultsUsed As Object, ByVal startLine As Long, ByVal beyondEnd As Boolean) As String
    Dim d As Object, emptyHits As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = True
    d("found") = False
    d("component") = component
    d("reason") = "no_more_hits"
    d("noMoreHits") = True
    Set emptyHits = New Collection
    Set d("hits") = emptyHits
    d("count") = 0
    Set d("requested") = requested
    d("countOfLines") = countOfLines
    d("searched") = False
    d("startLine") = startLine
    d("positionFieldsAreRequest") = True
    Set d("coords") = Owner.CoordsDictOfName(component, startLine, "ide")
    d("lineSpace") = "ide"
    If beyondEnd Then
        d("note") = "start_line=" & CStr(startLine) & " is past CountOfLines (" & CStr(countOfLines) & "): there are no more hits after the previous one. NOTHING was searched and the search does NOT wrap back to line 1 - this is the normal end of a pagination walk."
    Else
        d("note") = "there are no hits at or after start_line=" & CStr(startLine) & " in this window; the search does NOT wrap back to line 1."
    End If
    Set d("defaultsUsed") = defaultsUsed
    FindNoHitsJson = JsonStringify(d)
End Function

Public Function FindOpenProject(Owner As CodeTools, ByVal full As String) As VBIDE.VBProject
    Dim pj As VBIDE.VBProject
    If Len(full) = 0 Then Exit Function
    On Error Resume Next
    For Each pj In Owner.m_vbe.VBProjects
        If Len(pj.FileName) > 0 Then
            If StrComp(pj.FileName, full, vbTextCompare) = 0 Then
                Set FindOpenProject = pj
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next

    ' Build 2: the same file can be spelled differently (8.3 short name, forward
    ' slashes, different case, a directory alias). The exact-match pass above is
    ' unchanged; this second pass asks a canonical comparison instead.
    On Error Resume Next
    For Each pj In Owner.m_vbe.VBProjects
        If Len(pj.FileName) > 0 Then
            If Owner.SameProjectFile(pj.FileName, full) Then
                Set FindOpenProject = pj
                On Error GoTo 0
                Exit Function
            End If
        End If
    Next
    On Error GoTo 0
End Function

Public Function FindProject(Owner As CodeTools, ByVal name As String) As VBIDE.VBProject
    Dim p As VBIDE.VBProject
    For Each p In Owner.m_vbe.VBProjects
        If LCase$(p.Name) = LCase$(name) Then
            Set FindProject = p
            Exit Function
        End If
    Next
    Err.Raise vbObjectError + 46, "CodeTools", "Project not found: " & name
End Function

Public Function FindWindow(Owner As CodeTools, ByVal window As String) As VBIDE.Window
    Dim w As VBIDE.Window, idx As Long, want As Long
    If Len(Trim$(window)) = 0 Then Exit Function
    If IsNumeric(window) Then
        On Error Resume Next
        idx = CLng(Val(window))
        Set w = Owner.m_vbe.Windows.Item(idx)
        Err.Clear
        On Error GoTo 0
        If Not w Is Nothing Then
            Set FindWindow = w
            Exit Function
        End If
    End If
    want = Owner.WindowTypeOfName(window)
    If want >= 0 Then
        On Error Resume Next
        For Each w In Owner.m_vbe.Windows
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
    For Each w In Owner.m_vbe.Windows
        If InStr(1, w.Caption, window, vbTextCompare) > 0 Then
            Set FindWindow = w
            Exit For
        End If
    Next
    Err.Clear
    On Error GoTo 0
End Function

Public Function FindWindowDict(Owner As CodeTools, ByVal sLine As Long, ByVal sCol As Long, ByVal eLine As Long, ByVal eCol As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("startLine") = sLine
    d("startColumn") = sCol
    d("endLine") = eLine
    d("endColumn") = eCol
    Set FindWindowDict = d
End Function

Public Function FindWindowReject(Owner As CodeTools, ByVal component As String, ByVal reason As String, ByVal detail As String, ByVal requested As Object, ByVal countOfLines As Long, ByVal defaultsUsed As Object) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("found") = False
    d("reason") = reason
    d("component") = component
    Set d("requested") = requested
    d("countOfLines") = countOfLines
    Set d("coords") = Owner.CoordsDictOfName(component, 1, "ide")
    d("lineSpace") = "ide"
    d("lineSpaceHint") = "countOfLines is an IDE code-pane line count; get_component_code's numbered output is DISK lines: ideLine = diskLine - coords.headerOffset"
    Set d("defaultsUsed") = defaultsUsed
    d("detail") = detail
    d("note") = "nothing was searched and the component was NOT modified (the IDE Find engine was not called)"
    FindWindowReject = JsonStringify(d)
End Function

