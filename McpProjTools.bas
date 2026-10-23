Attribute VB_Name = "McpProjTools"
Option Explicit

' =======================================================================
' C-3 facade batch: McpProjTools (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Private Const VBP_ORDER_MAX_CHANGES As Long = 50
Private Const VBP_ORDER_MAX_LINES As Long = 400

Public Function ProjectControlParts(Owner As CodeTools) As Collection
    Dim col As New Collection
    Set ProjectControlParts = col
    On Error Resume Next
    Dim p As String
    p = McpCls.ProjectFileNameSafe(Owner)
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

Public Function ProjectDir(Owner As CodeTools) As String
    On Error GoTo np
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    If Len(p.FileName) = 0 Then Exit Function
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ProjectDir = fso.GetParentFolderName(p.FileName)
    Exit Function
np:
    ProjectDir = ""
End Function

Public Function ProjectDirWritesKilled(Owner As CodeTools) As Boolean
    Dim v As String
    v = LCase$(Trim$(Environ$("VB6MCP_ALLOW_PROJECTDIR_WRITES")))
    ProjectDirWritesKilled = (v = "0" Or v = "false" Or v = "no" Or v = "off")
End Function

Public Function ProjectDirtyScan(Owner As CodeTools, ByVal p As VBIDE.VBProject) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("dirtyCount") = 0
    d("dirtyNames") = ""
    d("dirtyUnknown") = False
    d("projectSaved") = True
    d("dirtyProject") = False
    If p Is Nothing Then
        ' no IDE surface: report the flags as unreadable, never as clean
        d("dirtyUnknown") = True
        Set ProjectDirtyScan = d
        Exit Function
    End If
    Dim names As String
    Dim cnt As Long
    Dim unk As Boolean
    Dim savedFlag As Boolean
    names = ""
    cnt = 0
    unk = False
    savedFlag = True
    On Error Resume Next
    Err.Clear
    savedFlag = p.Saved
    If Err.Number <> 0 Then
        unk = True
        Err.Clear
    End If
    On Error GoTo 0
    If Not savedFlag Then cnt = cnt + 1
    Dim comps As Object
    On Error Resume Next
    Err.Clear
    Set comps = p.VBComponents
    If Err.Number <> 0 Then
        unk = True
        Err.Clear
        Set comps = Nothing
    End If
    On Error GoTo 0
    If Not comps Is Nothing Then
        Dim comp As Object
        Dim cDirty As Boolean
        Dim cKnown As Boolean
        For Each comp In comps
            cDirty = False
            cKnown = False
            On Error Resume Next
            Err.Clear
            cDirty = comp.IsDirty
            If Err.Number = 0 Then cKnown = True Else Err.Clear
            On Error GoTo 0
            If Not cKnown Then
                unk = True
            ElseIf cDirty Then
                cnt = cnt + 1
                If Len(names) > 0 Then names = names & ","
                names = names & comp.Name
            End If
        Next comp
    End If
    d("dirtyCount") = cnt
    d("dirtyNames") = names
    d("dirtyUnknown") = unk
    d("projectSaved") = savedFlag
    d("dirtyProject") = Not savedFlag
    Set ProjectDirtyScan = d
End Function

Public Function ProjectHasUnsavedWork(Owner As CodeTools, ByVal projSaved As Boolean, ByVal savedReadable As Boolean, ByVal p As VBIDE.VBProject, ByRef info As Object) As Boolean
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

Public Function ProjectHostGuard(Owner As CodeTools, ByVal toolName As String) As String
    ProjectHostGuard = ""
    On Error Resume Next
    Dim ss As String
    Dim bs As Boolean
    Dim st As String
    Dim eb As Long
    ss = ""
    bs = False
    st = Owner.GetIdeStateEx(ss, bs)
    eb = IdeEbMode()
    If LCase$(st) = "design" And eb = 0 Then
        Err.Clear
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "ide_not_in_design"
    d("tool") = toolName
    d("state") = st
    d("stateSource") = ss
    d("ebMode") = eb
    d("commandBarsSuspended") = bs
    d("sideEffects") = "none: refused BEFORE any VBProjects host call"
    d("howToProceed") = "call control_execution{end} (or stop the program) until ide_state reports design, then retry; this tool never ends your program for you"
    d("note") = "a synchronous VBProjects call from break/run can wedge the IDE and has crashed it (measured)"
    ProjectHostGuard = JsonStringify(d)
    Err.Clear
End Function

Public Function ProjectListArr(Owner As CodeTools) As Object
    Dim c As Object
    Set c = New Collection
    Dim p As VBIDE.VBProject
    For Each p In Owner.m_vbe.VBProjects
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        d("name") = p.Name
        d("fileName") = p.FileName
        d("type") = ProjectTypeName(Owner, p)
        d("isStartProject") = False
        On Error Resume Next
        If Owner.m_vbe.VBProjects.StartProject Is p Then d("isStartProject") = True
        On Error GoTo 0
        c.Add d
    Next
    Set ProjectListArr = c
End Function

Public Function ProjectNameSafe(Owner As CodeTools) As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    ProjectNameSafe = p.Name
End Function

Public Function ProjectPathInfo(Owner As CodeTools, p As VBIDE.VBProject) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim opened As String, real As String, tier As Long, errText As String
    Dim aliasOn As Boolean
    d("projectPath") = ""
    d("realProjectPath") = ""
    d("projectPathIsAlias") = False
    d("realPathTier") = 0
    d("realPathError") = ""
    d("warning") = ""
    d("warningDetail") = ""
    d("refsFromAlias") = False
    d("refsBase") = ""
    If p Is Nothing Then
        Set ProjectPathInfo = d
        Exit Function
    End If
    aliasOn = Owner.RefsFromAliasOn()
    d("refsFromAlias") = aliasOn
    d("refsBase") = "canonical"
    If aliasOn Then d("refsBase") = "alias"
    On Error Resume Next
    opened = p.FileName
    If Err.Number <> 0 Then
        opened = ""
        Err.Clear
    End If
    On Error GoTo 0
    d("projectPath") = opened
    If Len(opened) = 0 Then
        Set ProjectPathInfo = d
        Exit Function
    End If
    real = Util.UtilRealPathTiers(opened, tier, errText)
    d("realPathTier") = tier
    d("realPathError") = errText
    If Len(real) = 0 Then
        d("warning") = "the canonical path of the project file could not be resolved (" & errText & "); relative references keep the directory the project was opened from"
        Set ProjectPathInfo = d
        Exit Function
    End If
    d("realProjectPath") = real
    If StrComp(LCase$(Replace(real, "/", "\")), LCase$(Replace(opened, "/", "\")), vbTextCompare) <> 0 Then
        d("projectPathIsAlias") = True
        d("warning") = "relative references are computed from the canonical path, not the link alias"
        d("warningDetail") = "the .vbp was opened as " & opened & " which resolves to " & real
    End If
    Set ProjectPathInfo = d
End Function

Public Function ProjectStartModeText(Owner As CodeTools, ByVal p As VBIDE.VBProject) As String
    ProjectStartModeText = ""
    On Error Resume Next
    Dim v As Variant
    v = p.Properties("StartMode")
    If Err.Number = 0 Then
        ProjectStartModeText = CStr(v)
    Else
        Err.Clear
    End If
End Function

Public Function ProjectTypeName(Owner As CodeTools, ByVal p As VBIDE.VBProject) As String
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

Public Function ProjectVbpText(Owner As CodeTools) As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    If p Is Nothing Then Exit Function
    ProjectVbpText = ProjShip.ReadTextFile(p.FileName)
End Function

Public Function VbpComponentKey(Owner As CodeTools, ByVal key As String) As Boolean
    VbpComponentKey = False
    Select Case key
        Case "form", "module", "class", "usercontrol", "propertypage", "userdocument", "designer"
            VbpComponentKey = True
    End Select
End Function

Public Function VbpDiskVsMemory(Owner As CodeTools, ByVal path As String, ByVal p As VBIDE.VBProject) As Object
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
            If VbpComponentKey(Owner, key) Then
                sp = InStrRev(val, ";")
                If sp > 0 Then val = Trim$(Mid$(val, sp + 1))
                If Len(val) > 0 Then diskSet(LCase$(Owner.BaseNameOf(val))) = True
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
                If Len(fn) > 0 Then memSet(LCase$(Owner.BaseNameOf(fn))) = True
            End If
            Err.Clear
        Next
    End If
    d("memoryProjectType") = memType
    d("memoryComponentCount") = memSet.Count
    d("typeMatches") = (Owner.NormalizeProjectTypeToken(diskType) = Owner.NormalizeProjectTypeToken(memType))
    Dim k As Variant
    For Each k In diskSet.Keys
        If Not memSet.Exists(CStr(k)) Then d("componentsOnlyOnDisk").Add CStr(k)
    Next
    For Each k In memSet.Keys
        If Not diskSet.Exists(CStr(k)) Then d("componentsOnlyInMemory").Add CStr(k)
    Next
    Owner.AddScalarDiff d, "ExeName32 vs BuildFileName", diskExe, memExe
    Owner.AddScalarDiff d, "Startup vs StartUpObject", diskStartup, memStartup
    d("matches") = (d("componentsOnlyOnDisk").Count = 0 And d("componentsOnlyInMemory").Count = 0 And CBool(d("typeMatches")) And d("decisiveScalarDiffs") = 0)
    If Err.Number <> 0 Then Err.Clear
End Function

Public Function VbpLineKey(Owner As CodeTools, ByVal line As String) As String
    Dim p As Long
    p = InStr(line, "=")
    If p > 1 Then
        VbpLineKey = Trim$(Left$(line, p - 1))
    Else
        VbpLineKey = Trim$(line)
    End If
End Function

Public Function VbpMissingProps(Owner As CodeTools, ByVal vbpPath As String, ByVal p As VBIDE.VBProject) As String
    VbpMissingProps = ""
    If Len(vbpPath) = 0 Then Exit Function
    Dim text As String
    On Error Resume Next
    text = Build.ReadLogAnsi(vbpPath)
    If Err.Number <> 0 Then
        Err.Clear
        VbpMissingProps = "<project file unreadable>"
        Exit Function
    End If
    Dim miss As String, v As String
    miss = ""
    v = p.Description
    If Err.Number <> 0 Then
        Err.Clear
        v = ""
    End If
    If Len(v) > 0 Then
        If InStr(1, text, "Description=" & Chr$(34) & v & Chr$(34), vbTextCompare) = 0 Then miss = "Description"
    End If
    v = p.HelpFile
    If Err.Number <> 0 Then
        Err.Clear
        v = ""
    End If
    If Len(v) > 0 Then
        If InStr(1, text, "HelpFile=" & Chr$(34) & v & Chr$(34), vbTextCompare) = 0 Then
            If Len(miss) > 0 Then miss = miss & ","
            miss = miss & "HelpFile"
        End If
    End If
    VbpMissingProps = miss
End Function

Public Function VbpOrderDelta(Owner As CodeTools, ByRef beforeSeq() As String, ByRef beforeIdx() As Long, ByRef afterSeq() As String, ByRef afterIdx() As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim changes As Object
    Set changes = New Collection
    d("preserved") = True
    d("sameSet") = VbpSameLineSet(Owner, beforeSeq, afterSeq)
    Set d("changes") = changes
    If UBound(beforeSeq) <> UBound(afterSeq) Then
        d("preserved") = False
        Set VbpOrderDelta = d
        Exit Function
    End If
    Dim used() As Boolean
    ReDim used(0 To UBound(afterSeq))
    Dim i As Long, j As Long, hit As Long, s As String
    For i = 0 To UBound(beforeSeq)
        hit = -1
        For j = 0 To UBound(afterSeq)
            If Not used(j) Then
                If afterSeq(j) = beforeSeq(i) Then
                    used(j) = True
                    hit = j
                    Exit For
                End If
            End If
        Next
        If hit < 0 Then
            d("preserved") = False
        ElseIf hit <> i Then
            d("preserved") = False
            If changes.Count < VBP_ORDER_MAX_CHANGES Then
                Dim c As Object
                Set c = CreateObject("Scripting.Dictionary")
                s = beforeSeq(i)
                c("key") = VbpLineKey(Owner, s)
                c("line") = s
                c("beforeLine") = beforeIdx(i) + 1
                c("afterLine") = afterIdx(hit) + 1
                c("beforeOrder") = i + 1
                c("afterOrder") = hit + 1
                changes.Add c
            End If
        End If
    Next
    Set VbpOrderDelta = d
End Function

Public Function VbpOrderSeq(Owner As CodeTools, ByRef lines() As String, ByRef outIdx() As Long) As String()
    Dim i As Long, n As Long, s As String
    Dim tmp() As String
    ReDim tmp(0 To UBound(lines))
    ReDim outIdx(0 To UBound(lines))
    n = 0
    For i = 0 To UBound(lines)
        s = Trim$(lines(i))
        If Len(s) > 0 Then
            If Left$(s, 1) <> "'" Then
                tmp(n) = lines(i)
                outIdx(n) = i
                n = n + 1
            End If
        End If
    Next
    If n = 0 Then
        ReDim tmp(0 To 0)
        ReDim outIdx(0 To 0)
    Else
        ReDim Preserve tmp(0 To n - 1)
        ReDim Preserve outIdx(0 To n - 1)
    End If
    VbpOrderSeq = tmp
End Function

Public Function VbpRefLineWith(Owner As CodeTools, ByVal line As String, ByVal newPath As String) As String
    Dim parts() As String, out As String, i As Long
    parts = Split(Trim$(line), "#", 5)
    out = parts(0)
    For i = 1 To UBound(parts)
        out = out & "#"
        If i = 3 Then out = out & newPath Else out = out & parts(i)
    Next
    VbpRefLineWith = out
End Function

Public Function VbpRefPathOf(Owner As CodeTools, ByVal line As String, ByRef refName As String) As String
    Dim t As String, parts() As String
    VbpRefPathOf = ""
    refName = ""
    t = Trim$(line)
    If Len(t) < 10 Then Exit Function
    If StrComp(Left$(t, 10), "Reference=", vbTextCompare) <> 0 Then Exit Function
    parts = Split(t, "#", 5)
    If UBound(parts) < 3 Then Exit Function
    VbpRefPathOf = Trim$(parts(3))
    If UBound(parts) >= 4 Then refName = Trim$(parts(4))
End Function

Public Sub VbpRefReportOnly(Owner As CodeTools, ByVal text As String, ByRef info As Object)
    Dim lines() As String, i As Long, ln As String, refPath As String, refName As String
    lines = McpCls.TextLines(Owner, text)
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 10 Then
            If StrComp(Left$(ln, 10), "Reference=", vbTextCompare) = 0 Then
                refPath = VbpRefPathOf(Owner, ln, refName)
                If Len(refPath) > 0 Then
                    Owner.RefsSkip info, refPath, "refs_from_alias", "the legacy switch is ON: relative references are NOT rewritten (A/B mode)"
                End If
            End If
        End If
    Next
End Sub

Public Function VbpRefRewriteText(Owner As CodeTools, ByVal text As String, ByVal aliasDir As String, ByVal realDir As String, ByVal projPaths As Object, ByRef info As Object) As String
    Dim lines() As String, eol As String, i As Long, ln As String
    Dim refPath As String, refName As String, candA As String, candB As String
    Dim realTarget As String, newPath As String, verifyPath As String
    Dim wrote As Boolean
    VbpRefRewriteText = text
    If Len(text) = 0 Then Exit Function
    If CBool(projPaths("refsFromAlias")) Then
        VbpRefReportOnly Owner, text, info
        Exit Function
    End If
    If Len(aliasDir) = 0 Or Len(realDir) = 0 Then Exit Function
    lines = McpCls.TextLines(Owner, text)
    eol = vbCrLf
    If InStr(text, vbCrLf) = 0 Then eol = vbLf
    wrote = False
    For i = 0 To UBound(lines)
        ln = lines(i)
        If Len(ln) > 10 Then
            If StrComp(Left$(ln, 10), "Reference=", vbTextCompare) = 0 Then
                refPath = VbpRefPathOf(Owner, ln, refName)
                If Len(refPath) > 0 Then
                    If InStr(refPath, "$(") > 0 Then
                        Owner.RefsSkip info, refPath, "environment_placeholder", "the path carries an environment placeholder, so it is never rewritten here"
                    ElseIf Not Util.UtilPathIsRelative(refPath) Then
                        Owner.RefsSkip info, refPath, "absolute_or_unc", "the path is absolute or UNC, so it does not depend on the project directory"
                    Else
                        candA = Owner.CtlJoinPath(aliasDir, refPath)
                        candB = Owner.CtlJoinPath(realDir, refPath)
                        realTarget = ""
                        If Owner.CtlFileExists(candA) Then realTarget = Util.UtilRealPath(candA)
                        If Len(realTarget) = 0 Then
                            If Owner.CtlFileExists(candB) Then realTarget = Util.UtilRealPath(candB)
                        End If
                        If Len(realTarget) = 0 Then
                            info("unresolvedCount") = CLng(info("unresolvedCount")) + 1
                            Owner.RefsSkip info, refPath, "unresolved", "neither " & candA & " nor " & candB & " names an existing file, so no canonical value is guessed"
                        Else
                            newPath = Util.UtilRelPath(realDir, realTarget)
                            If Len(newPath) = 0 Then
                                info("unresolvedCount") = CLng(info("unresolvedCount")) + 1
                                Owner.RefsSkip info, refPath, "relpath_failed", "no relative path from " & realDir & " to " & realTarget
                            Else
                                verifyPath = Owner.CtlJoinPath(realDir, newPath)
                                If Not Owner.CtlFileExists(verifyPath) Then
                                    info("unresolvedCount") = CLng(info("unresolvedCount")) + 1
                                    Owner.RefsSkip info, refPath, "verify_failed", "the rewritten path " & verifyPath & " does not name an existing file"
                                ElseIf StrComp(Util.UtilRealPath(verifyPath), realTarget, vbTextCompare) <> 0 Then
                                    info("unresolvedCount") = CLng(info("unresolvedCount")) + 1
                                    Owner.RefsSkip info, refPath, "verify_failed", "the rewritten path " & verifyPath & " resolves to " & Util.UtilRealPath(verifyPath) & ", not to " & realTarget
                                Else
                                    If StrComp(newPath, refPath, vbTextCompare) = 0 Then
                                        ' the same path, only the SPELLING differs (case / 8.3):
                                        ' keep the file's own spelling so a plain save stays
                                        ' byte-stable instead of churning the line.
                                        Owner.RefsRecord info, refName, refPath, refPath, False, realDir, realTarget
                                    Else
                                        lines(i) = VbpRefLineWith(Owner, ln, newPath)
                                        wrote = True
                                        Owner.RefsRecord info, refName, refPath, newPath, True, realDir, realTarget
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    If wrote Then VbpRefRewriteText = Join(lines, eol)
End Function

Public Function VbpRefsRepairOnDisk(Owner As CodeTools, ByVal vbpPath As String, ByVal projPaths As Object, ByRef info As Object) As Long
    Dim aliasDir As String, realDir As String, text As String, outText As String
    VbpRefsRepairOnDisk = 0
    If Len(vbpPath) = 0 Then Exit Function
    If CBool(projPaths("refsFromAlias")) Then Exit Function
    aliasDir = Owner.DirOfPath(CStr(projPaths("projectPath")))
    realDir = Owner.DirOfPath(CStr(projPaths("realProjectPath")))
    If Len(aliasDir) = 0 Or Len(realDir) = 0 Then Exit Function
    text = ProjShip.ReadTextFile(vbpPath)
    If Len(text) = 0 Then Exit Function
    outText = VbpRefRewriteText(Owner, text, aliasDir, realDir, projPaths, info)
    If outText = text Then Exit Function
    If ProjShip.WriteTextFile(vbpPath, outText) Then VbpRefsRepairOnDisk = 1
End Function

Public Function VbpRestoreOrder(Owner As CodeTools, ByVal afterText As String, ByRef beforeSeq() As String, ByRef restoredText As String) As Boolean
    VbpRestoreOrder = False
    restoredText = ""
    Dim lines() As String
    lines = McpCls.TextLines(Owner, afterText)
    Dim seq() As String, idx() As Long
    seq = VbpOrderSeq(Owner, lines, idx)
    If Not VbpSameLineSet(Owner, beforeSeq, seq) Then Exit Function
    If UBound(idx) <> UBound(beforeSeq) Then Exit Function
    Dim k As Long
    For k = 0 To UBound(beforeSeq)
        lines(idx(k)) = beforeSeq(k)
    Next
    Dim eol As String
    eol = vbCrLf
    If InStr(afterText, vbCrLf) = 0 Then eol = vbLf
    ' Split/Join is an exact round trip: a trailing terminator becomes a trailing empty
    ' element and Join puts it back, so no extra terminator may be appended here.
    restoredText = Join(lines, eol)
    VbpRestoreOrder = True
End Function

Public Function VbpSameLineSet(Owner As CodeTools, ByRef a() As String, ByRef b() As String) As Boolean
    VbpSameLineSet = False
    Dim i As Long, j As Long
    If UBound(a) <> UBound(b) Then Exit Function
    If UBound(a) = 0 Then
        VbpSameLineSet = (Len(a(0)) = 0 And Len(b(0)) = 0)
        Exit Function
    End If
    Dim used() As Boolean
    ReDim used(0 To UBound(b))
    Dim found As Boolean
    For i = 0 To UBound(a)
        found = False
        For j = 0 To UBound(b)
            If Not used(j) Then
                If b(j) = a(i) Then
                    used(j) = True
                    found = True
                    Exit For
                End If
            End If
        Next
        If Not found Then Exit Function
    Next
    VbpSameLineSet = True
End Function

Public Function VbpSeqJsonArray(Owner As CodeTools, ByRef seq() As String) As Object
    Dim out As Object
    Set out = New Collection
    Dim i As Long
    For i = 0 To UBound(seq)
        If i >= VBP_ORDER_MAX_LINES Then Exit For
        out.Add seq(i)
    Next
    Set VbpSeqJsonArray = out
End Function

Public Function VbpSettingValue(Owner As CodeTools, ByVal text As String, ByVal settingName As String) As String
    VbpSettingValue = ""
    If Len(text) = 0 Then Exit Function
    Dim lines() As String, i As Long
    lines = McpCls.TextLines(Owner, text)
    For i = 0 To UBound(lines)
        If Left$(lines(i), Len(settingName) + 1) = settingName & "=" Then
            VbpSettingValue = Trim$(Mid$(lines(i), Len(settingName) + 2))
            Exit Function
        End If
    Next
End Function

