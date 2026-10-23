Attribute VB_Name = "McpBpPane"
Option Explicit

' =======================================================================
' C-3 facade batch: McpBpPane (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Public Function BpAnchorPaneLine(Owner As CodeTools, ByVal comp As String, ByVal anchor As String) As Long
    BpAnchorPaneLine = 0
    On Error Resume Next
    Dim c As VBIDE.VBComponent
    Set c = McpFindTools.FindComponent(Owner, comp)
    If c Is Nothing Then Err.Clear : Exit Function
    Dim cm As VBIDE.CodeModule
    Set cm = c.CodeModule
    If cm Is Nothing Then Err.Clear : Exit Function
    Dim want As String
    want = LCase$(Trim$(anchor))
    Dim i As Long
    For i = 1 To cm.CountOfLines
        If LCase$(Trim$(cm.Lines(i, 1))) = want Then
            BpAnchorPaneLine = i
            Exit Function
        End If
    Next
    Err.Clear
End Function

Public Function BpClassifyPaneLine(Owner As CodeTools, ByVal comp As String, ByVal paneLine As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("component") = comp
    d("paneLine") = paneLine
    d("lineClass") = "unknown"
    d("breakable") = False
    d("reason") = ""
    d("lineText") = ""
    Dim near As Object
    Set near = New Collection
    Set d("nearestBreakable") = near
    On Error Resume Next
    Dim c As VBIDE.VBComponent
    Set c = McpFindTools.FindComponent(Owner, comp)
    If c Is Nothing Then
        d("lineClass") = "component_not_found"
        d("reason") = "the component was not found in the open project"
        Set BpClassifyPaneLine = d
        Err.Clear
        Exit Function
    End If
    Dim cm As VBIDE.CodeModule
    Set cm = c.CodeModule
    If cm Is Nothing Then
        d("lineClass") = "no_code_module"
        d("reason") = "the component has no code module (a designer-only component)"
        Set BpClassifyPaneLine = d
        Err.Clear
        Exit Function
    End If
    Dim cls As Object
    Set cls = BpLineClassOf(Owner, cm, paneLine)
    d("lineClass") = cls("lineClass")
    d("lineText") = cls("text")
    d("breakable") = cls("breakable")
    If Not CBool(cls("breakable")) Then d("reason") = "line_not_breakable: " & CStr(cls("lineClass"))
    Dim step As Long
    Dim other As Object
    For step = 1 To 5
        If paneLine - step >= 1 Then
            Set other = BpLineClassOf(Owner, cm, paneLine - step)
            If CBool(other("breakable")) Then near.Add other
        End If
    Next
    For step = 1 To 5
        If paneLine + step <= cm.CountOfLines Then
            Set other = BpLineClassOf(Owner, cm, paneLine + step)
            If CBool(other("breakable")) Then near.Add other
        End If
    Next
    Err.Clear
    Set BpClassifyPaneLine = d
End Function

Public Function BpFindAt(Owner As CodeTools, ByVal comp As String, ByVal paneLine As Long) As Object
    Set BpFindAt = Nothing
    Dim reg As Object
    Set reg = McpBpReg.BpRegistry(Owner)
    Dim k As Variant
    For Each k In reg.Keys
        Dim rec As Object
        Set rec = reg(k)
        If StrComp(CStr(rec("component")), comp, vbTextCompare) = 0 Then
            If CLng(rec("paneLine")) = paneLine Then
                Set BpFindAt = rec
                Exit Function
            End If
        End If
    Next
End Function

Public Function BpFindInjected(Owner As CodeTools, ByVal id As String) As Object
    Set BpFindInjected = Nothing
    Dim inj As Object
    Set inj = BpInjectedList(Owner)
    Dim i As Long
    For i = 1 To inj.Count
        Dim e As Object
        Set e = inj(i)
        If StrComp(JsonGetStr(e, "id", ""), id, vbTextCompare) = 0 Then
            Set BpFindInjected = e
            Exit Function
        End If
    Next
End Function

Public Function BpFindNativeRecord(Owner As CodeTools, ByVal comp As String, ByVal paneLine As Long) As Object
    Set BpFindNativeRecord = Nothing
    On Error Resume Next
    Dim reg As Object
    Set reg = McpBpReg.BpRegistry(Owner)
    Dim k As Variant
    For Each k In reg.Keys
        Dim rec As Object
        Set rec = reg(k)
        If LCase$(JsonGetStr(rec, "mode", "")) = "native" Then
            If StrComp(JsonGetStr(rec, "component", ""), comp, vbTextCompare) = 0 Then
                If CLng(JsonGetLong(rec, "paneLine", 0)) = paneLine Then
                    Set BpFindNativeRecord = rec
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
    Next
    Dim disk As Object
    Set disk = McpBpReg.BpRegLoad(Owner)
    Dim i As Long
    For i = 1 To disk.Count
        Dim e As Object
        Set e = disk(i)
        If LCase$(JsonGetStr(e, "mode", "")) = "native" Then
            If StrComp(JsonGetStr(e, "component", ""), comp, vbTextCompare) = 0 Then
                If CLng(JsonGetLong(e, "paneLine", 0)) = paneLine Then
                    Dim r2 As Object
                    Set r2 = CreateObject("Scripting.Dictionary")
                    r2("id") = JsonGetStr(e, "id", "")
                    r2("component") = comp
                    r2("paneLine") = paneLine
                    r2("diskLine") = JsonGetLong(e, "diskLine", paneLine)
                    r2("mode") = "native"
                    r2("enabled") = JsonGetBool(e, "enabled", True)
                    r2("paneShaAtSet") = JsonGetStr(e, "paneShaAtSet", "")
                    r2("setAt") = JsonGetStr(e, "setAt", "")
                    r2("from") = "disk"
                    r2("origin") = "disk"
                    r2("confidence") = "tool_claim_only"
                    r2("hitCount") = 0
                    r2("hitCountTarget") = 1
                    reg.Add CStr(r2("id")), r2
                    Set BpFindNativeRecord = r2
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
    Next
    Err.Clear
End Function

Public Function BpHitsOfAnchor(Owner As CodeTools, ByVal comp As String, ByVal anchor As String) As Object
    Set BpHitsOfAnchor = New Collection
    On Error Resume Next
    Dim c As VBIDE.VBComponent
    Set c = McpFindTools.FindComponent(Owner, comp)
    If c Is Nothing Then Err.Clear : Exit Function
    Dim cm As VBIDE.CodeModule
    Set cm = c.CodeModule
    If cm Is Nothing Then Err.Clear : Exit Function
    Dim want As String
    want = LCase$(Trim$(anchor))
    If Len(want) = 0 Then Err.Clear : Exit Function
    Dim i As Long
    For i = 1 To cm.CountOfLines
        If LCase$(Trim$(cm.Lines(i, 1))) = want Then
            Dim h As Object
            Set h = CreateObject("Scripting.Dictionary")
            h("paneLine") = i
            h("text") = cm.Lines(i, 1)
            Dim pr As String
            pr = ""
            On Error Resume Next
            pr = cm.ProcOfLine(i, 0)
            If Err.Number <> 0 Then pr = ""
            Err.Clear
            On Error GoTo 0
            h("proc") = pr
            BpHitsOfAnchor.Add h
        End If
    Next
    Err.Clear
End Function

Public Function BpInjectedList(Owner As CodeTools) As Object
    On Error Resume Next
    Dim out As Object
    Set out = New Collection
    Dim rj As String
    rj = Owner.ToolListBreakpoints()
    Dim v As Variant
    If Not JsonParse(rj, v) Then Set BpInjectedList = out : Exit Function
    Dim d As Object
    Set d = v
    If d Is Nothing Then Set BpInjectedList = out : Exit Function
    If Not d.Exists("injected") Then Set BpInjectedList = out : Exit Function
    ' JsonLib parses an array into a Collection that is STORED with Set; reading it back
    ' therefore needs Set as well (a plain Variant assignment raises and would look empty).
    Dim arr As Object
    Set arr = d("injected")
    If arr Is Nothing Then Set BpInjectedList = out : Exit Function
    Dim e As Variant
    For Each e In arr
        If IsObject(e) Then out.Add e
    Next
    Set BpInjectedList = out
End Function

Public Function BpIsLabelLine(Owner As CodeTools, ByVal t As String) As Boolean
    BpIsLabelLine = False
    If Len(t) < 2 Then Exit Function
    If Right$(t, 1) <> ":" Then Exit Function
    Dim i As Long, ch As String
    ch = Left$(t, 1)
    If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Or ch = "_") Then Exit Function
    For i = 2 To Len(t) - 1
        ch = Mid$(t, i, 1)
        If Not ((ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") _
                Or (ch >= "0" And ch <= "9") Or ch = "_") Then Exit Function
    Next
    BpIsLabelLine = True
End Function

Public Function BpLineClassOf(Owner As CodeTools, ByVal cm As VBIDE.CodeModule, ByVal ln As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("line") = ln
    d("lineClass") = "unknown"
    d("breakable") = False
    d("text") = ""
    On Error Resume Next
    If cm Is Nothing Then
        d("lineClass") = "no_code_module"
        Set BpLineClassOf = d
        Exit Function
    End If
    If ln < 1 Or ln > cm.CountOfLines Then
        d("lineClass") = "out_of_range"
        Set BpLineClassOf = d
        Exit Function
    End If
    Dim txt As String
    txt = cm.Lines(ln, 1)
    d("text") = txt
    Dim t As String
    t = Trim$(txt)
    Dim u As String
    u = UCase$(t)
    Dim prev As String
    prev = ""
    If ln > 1 Then prev = cm.Lines(ln - 1, 1)
    Dim procName As String
    procName = ""
    procName = cm.ProcOfLine(ln, 0)
    Err.Clear
    If Len(t) = 0 Then
        d("lineClass") = "blank_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Left$(t, 1) = "'" Then
        d("lineClass") = "comment_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If u = "REM" Or Left$(u, 4) = "REM " Then
        d("lineClass") = "comment_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Left$(u, 10) = "ATTRIBUTE " Then
        d("lineClass") = "attribute_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Right$(RTrim$(prev), 1) = "_" Then
        d("lineClass") = "continuation_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If u = "END SUB" Or u = "END FUNCTION" Or u = "END PROPERTY" Then
        d("lineClass") = "procedure_footer"
        Set BpLineClassOf = d
        Exit Function
    End If
    Dim hdr As String
    hdr = u
    Do While True
        If Left$(hdr, 7) = "PUBLIC " Then
            hdr = Mid$(hdr, 8)
        ElseIf Left$(hdr, 8) = "PRIVATE " Then
            hdr = Mid$(hdr, 9)
        ElseIf Left$(hdr, 7) = "FRIEND " Then
            hdr = Mid$(hdr, 8)
        ElseIf Left$(hdr, 7) = "STATIC " Then
            hdr = Mid$(hdr, 8)
        ElseIf Left$(hdr, 7) = "GLOBAL " Then
            hdr = Mid$(hdr, 8)
        Else
            Exit Do
        End If
    Loop
    If Left$(hdr, 4) = "SUB " Or Left$(hdr, 9) = "FUNCTION " Or Left$(hdr, 9) = "PROPERTY " Then
        d("lineClass") = "procedure_header"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Left$(t, 1) = "#" Then
        d("lineClass") = "compiler_directive"
        Set BpLineClassOf = d
        Exit Function
    End If
    If BpIsLabelLine(Owner, t) Then
        d("lineClass") = "label_line"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Len(procName) = 0 Then
        If Left$(u, 6) = "OPTION" Or Left$(u, 7) = "DECLARE" Or Left$(u, 3) = "DIM" _
           Or Left$(u, 5) = "CONST" Or Left$(u, 4) = "TYPE" Or Left$(u, 4) = "ENUM" _
           Or Left$(u, 5) = "EVENT" Or Left$(u, 10) = "IMPLEMENTS" Then
            d("lineClass") = "module_declaration"
            Set BpLineClassOf = d
            Exit Function
        End If
        d("lineClass") = "module_level_other"
        Set BpLineClassOf = d
        Exit Function
    End If
    If Left$(u, 4) = "DIM " Or Left$(u, 6) = "STATIC" And Left$(u, 7) = "STATIC " Then
        d("lineClass") = "local_decl"
        d("breakable") = True
        Set BpLineClassOf = d
        Exit Function
    End If
    d("lineClass") = "statement_line"
    d("breakable") = True
    Set BpLineClassOf = d
End Function

Public Function BpNewRecord(Owner As CodeTools, ByVal comp As String, ByVal paneLine As Long, ByVal mode As String, ByVal id As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("id") = id
    d("component") = comp
    d("paneLine") = paneLine
    d("mode") = mode
    d("enabled") = True
    d("setAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    d("hitCount") = 0
    d("hits") = 0
    d("paneShaAtSet") = BpPaneSha(Owner, comp)
    On Error Resume Next
    Dim cd As Object
    Set cd = Owner.CoordsDictOfName(comp, paneLine, "ide")
    If Err.Number <> 0 Then
        Err.Clear
        Set cd = CreateObject("Scripting.Dictionary")
    End If
    Set d("coords") = cd
    d("diskLine") = paneLine
    d("diskLineExact") = False
    d("mappingResolved") = False
    If Not cd Is Nothing Then
        If cd.Exists("diskLine") Then d("diskLine") = cd("diskLine")
        If cd.Exists("diskLineExact") Then d("diskLineExact") = cd("diskLineExact")
        If cd.Exists("mappingResolved") Then d("mappingResolved") = cd("mappingResolved")
    End If
    d("coordsNote") = "paneLine is the line an edit tool takes; diskLine comes from the SAME cumulative pane->disk alignment the edit tools report (coords.mappingResolved=true means exact). While the pane holds edits that are not on disk the alignment can be inexact and diskLineExact=false - the scalar headerOffset fallback is then NOT trusted."
    Set BpNewRecord = d
End Function

Public Function BpPaneSha(Owner As CodeTools, ByVal comp As String) As String
    On Error Resume Next
    BpPaneSha = ""
    Dim got As Boolean
    got = False
    Dim t As String
    t = Owner.PaneTextOfName(comp, got)
    BpPaneSha = Sha256.Sha256HexTextUtf8(t)
    Err.Clear
End Function

Public Function BpProjectDir(Owner As CodeTools) As String
    BpProjectDir = ""
    On Error Resume Next
    BpProjectDir = McpProjTools.ProjectDir(Owner)
    Err.Clear
End Function

Public Function BpRecOf(Owner As CodeTools, ByVal id As String) As Object
    Set BpRecOf = Nothing
    If Len(id) = 0 Then Exit Function
    Dim reg As Object
    Set reg = McpBpReg.BpRegistry(Owner)
    If reg.Exists(id) Then Set BpRecOf = reg(id)
End Function

Public Function BpRefusalBlock(Owner As CodeTools, ByVal cls As Object) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = "line_not_breakable"
    d("lineClass") = cls("lineClass")
    d("paneLine") = cls("paneLine")
    d("component") = cls("component")
    d("lineText") = cls("lineText")
    Set d("nearestBreakable") = cls("nearestBreakable")
    d("menuCommandSent") = False
    d("howToProceed") = "choose a statement line INSIDE a procedure body (the nearestBreakable list gives the closest legal pane lines); mode=injected also works and never opens a dialog"
    Set BpRefusalBlock = d
End Function

Public Function BpStartsWord(Owner As CodeTools, ByVal loLine As String, ByVal w As String) As Boolean
    If Len(loLine) <= Len(w) Then Exit Function
    If Left$(loLine, Len(w)) <> w Then Exit Function
    BpStartsWord = True
End Function

