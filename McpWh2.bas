Attribute VB_Name = "McpWh2"
Option Explicit

' =======================================================================
' C-3 facade batch: McpWh2 (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Public Sub Wh2AddVar(Owner As CodeTools, ByRef out As Object, ByRef seenNames As Object, ByVal decl As String, ByVal kind As String, ByVal srcLine As Long)
    On Error Resume Next
    Dim s As String
    s = Trim$(decl)
    ' strip modifiers and an optional "As <type>"
    Dim stripped As Boolean
    Do
        stripped = False
        If Len(s) >= 9 Then
            If StrComp(Left$(s, 8), "Optional ", vbTextCompare) = 0 Then
                s = Trim$(Mid$(s, 9))
                stripped = True
            End If
        End If
        If Not stripped Then
            If Len(s) >= 6 Then
                If StrComp(Left$(s, 5), "ByVal ", vbTextCompare) = 0 Then
                    s = Trim$(Mid$(s, 6))
                    stripped = True
                End If
            End If
        End If
        If Not stripped Then
            If Len(s) >= 6 Then
                If StrComp(Left$(s, 5), "ByRef ", vbTextCompare) = 0 Then
                    s = Trim$(Mid$(s, 6))
                    stripped = True
                End If
            End If
        End If
        If Not stripped Then
            If Len(s) >= 11 Then
                If StrComp(Left$(s, 10), "ParamArray ", vbTextCompare) = 0 Then
                    s = Trim$(Mid$(s, 11))
                    stripped = True
                End If
            End If
        End If
        If Not stripped Then Exit Do
    Loop
    Dim declaredType As String
    declaredType = ""
    Dim i As Long
    i = InStr(1, s, " As ", vbTextCompare)
    If i > 0 Then
        declaredType = Trim$(Mid$(s, i + 4))
        s = Trim$(Left$(s, i - 1))
    End If
    ' drop array parentheses and any default value
    Dim p As Long
    p = InStr(s, "(")
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    p = InStr(s, "=")
    If p > 0 Then s = Trim$(Left$(s, p - 1))
    If Len(s) = 0 Then Exit Sub
    If seenNames.Exists(s) Then Exit Sub
    seenNames.Add s, True
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("name") = s
    d("kind") = kind
    d("declaredType") = declaredType
    d("paneLine") = srcLine
    d("paneLineSpace") = "ide"
    d("nameSource") = "source declaration (" & kind & ")"
    out.Add d
    Err.Clear
End Sub

Public Function Wh2EvalRaw(Owner As CodeTools, ByVal expr As String) As Object
    On Error Resume Next
    Set Wh2EvalRaw = Nothing
    Dim rj As String
    rj = Owner.ToolImmediateEval(expr, "", 5000)
    If Len(rj) = 0 Then Exit Function
    Dim v As Variant
    If Not JsonParse(rj, v) Then Exit Function
    Dim d As Object
    Set d = v
    Set Wh2EvalRaw = d
    Err.Clear
End Function

Public Function Wh2EvalStr(Owner As CodeTools, ByVal expr As String, ByVal frameIndex As Long) As String
    On Error Resume Next
    Wh2EvalStr = ""
    Dim d As Object
    Set d = Wh2EvalRaw(Owner, expr)
    If d Is Nothing Then Exit Function
    If d.Exists("outputTrimmed") Then Wh2EvalStr = CStr(d("outputTrimmed"))
    If Len(Wh2EvalStr) = 0 Then
        If d.Exists("value") Then Wh2EvalStr = CStr(d("value"))
    End If
    Err.Clear
End Function

Public Function Wh2FrameContext(Owner As CodeTools, ByVal frameIndex As Long, ByRef diag As Object) As String
    Wh2FrameContext = ""
    diag("frameBinding") = "ideImmediateContext (the IDE's own current frame; NO engine-internal frame call is made)"
    diag("frameBindingNote") = "the VBA6 frame-set thunk is deliberately NOT used: it terminated the IDE in a measured run. Use vba6_probe{frame_set:true} explicitly if you accept that risk."
    If frameIndex <> 0 Then
        Wh2FrameContext = "frame_index " & CStr(frameIndex) & " is an OUTER frame: reaching it needs the engine-internal frame-set call, which is refused here because it can terminate the IDE (measured). Only frame 0 (the IDE's current frame) is available."
        Exit Function
    End If
    If Owner.GetIdeState() <> "break" Then Wh2FrameContext = "not_in_break"
End Function

Public Function Wh2Inspect(Owner As CodeTools, ByVal expr As String, ByVal frameIndex As Long, ByRef diag As Object, ByVal knownType As String) As Object
    Dim seen As Object
    Set seen = New Collection
    Set Wh2Inspect = Wh2InspectNode(Owner, expr, frameIndex, diag, 1, seen, knownType)
End Function

Public Function Wh2InspectNode(Owner As CodeTools, ByVal expr As String, ByVal frameIndex As Long, ByRef diag As Object, ByVal level As Long, ByRef seen As Object, ByVal knownType As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("depthLimit") = 2
    d("widthLimit") = 20
    d("truncatedDepth") = False
    d("truncatedWidth") = False
    d("cycleDetected") = False
    Dim seedItems As Object
    Set seedItems = New Collection
    Set d("items") = seedItems
    On Error Resume Next
    Dim tn As String
    tn = ""
    If Len(knownType) > 0 Then tn = knownType
    If Len(tn) = 0 Then tn = Wh2EvalStr(Owner, "TypeName(" & expr & ")", frameIndex)
    d("typeName") = tn
    If Wh2IsScalarType(Owner, tn) Then
        d("kind") = "scalar"
        d("isObject") = False
        d("note") = "scalar: no object probing, no extra evaluation"
        Set Wh2InspectNode = d
        Exit Function
    End If
    If InStr(1, tn, "()", vbTextCompare) > 0 Then
        d("kind") = "array"
        Dim lb As Long
        Dim ub As Long
        lb = CLng(Val(Wh2EvalStr(Owner, "LBound(" & expr & ")", frameIndex)))
        ub = CLng(Val(Wh2EvalStr(Owner, "UBound(" & expr & ")", frameIndex)))
        d("lowerBound") = lb
        d("upperBound") = ub
        Dim cnt As Long
        cnt = 0
        If ub >= lb Then cnt = ub - lb + 1
        d("count") = cnt
        Dim n As Long
        n = cnt
        If n > 20 Then
            n = 20
            d("truncatedWidth") = True
        End If
        Dim items As Object
        Set items = New Collection
        Dim i As Long
        For i = lb To lb + n - 1
            If level >= 2 Then
                Dim leaf As Object
                Set leaf = Wh2InspectNode(Owner, expr & "(" & CStr(i) & ")", frameIndex, diag, level + 1, seen, "")
                leaf("index") = i
                items.Add leaf
            Else
                Dim s As String
                s = Wh2EvalStr(Owner, "CStr(" & expr & "(" & CStr(i) & "))", frameIndex)
                Dim it As Object
                Set it = CreateObject("Scripting.Dictionary")
                it("index") = i
                it("value") = s
                items.Add it
            End If
        Next
        Set d("items") = items
    ElseIf StrComp(tn, "Dictionary", vbTextCompare) = 0 Then
        d("kind") = "dictionary"
        Dim dcnt As Long
        dcnt = CLng(Val(Wh2EvalStr(Owner, expr & ".Count", frameIndex)))
        d("count") = dcnt
        Dim dn As Long
        dn = dcnt
        If dn > 20 Then
            dn = 20
            d("truncatedWidth") = True
        End If
        Dim ditems As Object
        Set ditems = New Collection
        Dim j As Long
        For j = 0 To dn - 1
            Dim pair As Object
            Set pair = CreateObject("Scripting.Dictionary")
            pair("key") = Wh2EvalStr(Owner, "CStr(" & expr & ".Keys()(" & CStr(j) & "))", frameIndex)
            If level >= 2 Then
                pair("value") = Wh2EvalStr(Owner, "TypeName(" & expr & ".Items()(" & CStr(j) & "))", frameIndex)
                pair("valueKind") = "typeOnly (depth limit reached)"
                d("truncatedDepth") = True
            Else
                pair("value") = Wh2EvalStr(Owner, "CStr(" & expr & ".Items()(" & CStr(j) & "))", frameIndex)
            End If
            ditems.Add pair
        Next
        Set d("items") = ditems
    ElseIf Not Wh2IsScalarType(Owner, tn) Then
        d("kind") = "customObject"
        d("fieldsUnavailable") = True
        ' identity for cycle detection; VB6 has NO reflection, so fields cannot be enumerated
        Dim ptr As String
        ptr = Wh2EvalStr(Owner, "ObjPtr(" & expr & ")", frameIndex)
        d("objPtr") = ptr
        If Len(ptr) > 0 Then
            Dim k As Variant
            For Each k In seen
                If StrComp(CStr(k), ptr, vbTextCompare) = 0 Then d("cycleDetected") = True
            Next
            If Not CBool(d("cycleDetected")) Then seen.Add ptr
        End If
        d("note") = "VB6 has no reflection: a custom class exposes only its TypeName and ObjPtr; fields are NOT guessed"
    Else
        d("kind") = "scalar"
        Dim sv As String
        sv = Wh2EvalStr(Owner, "CStr(" & expr & ")", frameIndex)
        d("value") = sv
    End If
    Err.Clear
    Set Wh2InspectNode = d
End Function

Public Function Wh2IsScalarType(Owner As CodeTools, ByVal tn As String) As Boolean
    Dim t As String
    t = LCase$(Trim$(tn))
    Wh2IsScalarType = False
    If Len(t) = 0 Then Exit Function
    If InStr(1, t, "()", vbTextCompare) > 0 Then Exit Function
    Select Case t
        Case "byte", "boolean", "integer", "long", "single", "double", "currency", "date", "string", "variant", "empty", "null"
            Wh2IsScalarType = True
    End Select
End Function

Public Function Wh2NameAllowed(Owner As CodeTools, ByVal name As String) As Boolean
    Wh2NameAllowed = False
    If Owner.m_lastLocalNames Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To Owner.m_lastLocalNames.Count
        If StrComp(CStr(Owner.m_lastLocalNames(i)), name, vbTextCompare) = 0 Then
            Wh2NameAllowed = True
            Exit Function
        End If
    Next
End Function

Public Function Wh2SourceVars(Owner As CodeTools, ByVal moduleName As String, ByVal procName As String, ByRef root As Object) As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    Dim scan As Object
    Set scan = CreateObject("Scripting.Dictionary")
    scan("module") = moduleName
    scan("proc") = procName
    Set root("sourceScan") = scan
    If Len(moduleName) = 0 Then
        scan("reason") = "module_name_empty"
        Set Wh2SourceVars = out
        Exit Function
    End If
    If Len(procName) = 0 Then
        scan("reason") = "proc_name_empty"
        Set Wh2SourceVars = out
        Exit Function
    End If
    Dim vbp As VBIDE.VBProject
    Set vbp = Owner.GetProject()
    If vbp Is Nothing Then
        scan("reason") = "no_project"
        Set Wh2SourceVars = out
        Exit Function
    End If
    Dim comp As VBIDE.VBComponent
    Err.Clear
    Set comp = vbp.VBComponents.Item(moduleName)
    If Err.Number <> 0 Then
        scan("reason") = "component_not_found: " & Err.Description
        Err.Clear
        Set Wh2SourceVars = out
        Exit Function
    End If
    If comp Is Nothing Then
        scan("reason") = "component_not_found: VBComponents.Item returned Nothing"
        Set Wh2SourceVars = out
        Exit Function
    End If
    scan("component") = comp.Name
    Dim cm As VBIDE.CodeModule
    Err.Clear
    Set cm = comp.CodeModule
    If Err.Number <> 0 Then
        scan("reason") = "code_module_error: " & Err.Description
        Err.Clear
        Set Wh2SourceVars = out
        Exit Function
    End If
    If cm Is Nothing Then
        scan("reason") = "code_module_missing"
        Set Wh2SourceVars = out
        Exit Function
    End If
    Dim sl As Long
    Dim nl As Long
    sl = 0
    nl = 0
    sl = cm.ProcStartLine(procName, 0)
    scan("procStartLine") = sl
    nl = cm.ProcCountLines(procName, 0)
    scan("procCountLines") = nl
    If sl < 1 Or nl < 1 Then
        scan("reason") = "proc_range_unavailable"
        Set Wh2SourceVars = out
        Exit Function
    End If
    Dim txt As String
    txt = ""
    txt = cm.Lines(sl, nl)
    scan("textLen") = Len(txt)
    scan("codeModuleFound") = True
    Dim lines() As String
    lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim seenNames As Object
    Set seenNames = CreateObject("Scripting.Dictionary")
    ' 1) the parameter list (the signature may continue over several lines)
    Dim sig As String
    Dim k As Long
    Dim sigLine As Long
    sigLine = 0
    sig = ""
    For k = 0 To UBound(lines)
        If InStr(lines(k), "(") > 0 Then
            sigLine = k
            sig = lines(k)
            Do While InStr(sig, ")") = 0 And k < UBound(lines)
                k = k + 1
                sig = sig & " " & lines(k)
            Loop
            Exit For
        End If
    Next
    If InStr(sig, "(") > 0 Then
        Dim a As Long
        Dim b As Long
        a = InStr(sig, "(")
        b = InStrRev(sig, ")")
        If b > a Then
            Dim inner As String
            inner = Mid$(sig, a + 1, b - a - 1)
            If Len(Trim$(inner)) > 0 Then
                Dim parts() As String
                parts = Split(inner, ",")
                Dim pi As Long
                For pi = 0 To UBound(parts)
                    Wh2AddVar Owner, out, seenNames, parts(pi), "parameter", sl + sigLine
                Next
            End If
        End If
    End If
    ' 2) Dim / Static declarations in the body (comma-separated declarators supported)
    scan("parsedParams") = out.Count
    Dim li As Long
    For li = 0 To UBound(lines)
        Dim t As String
        t = Trim$(lines(li))
        If Len(t) > 4 Then
            If StrComp(Left$(t, 4), "Dim ", vbTextCompare) = 0 Or StrComp(Left$(t, 7), "Static ", vbTextCompare) = 0 Then
                Dim body As String
                If StrComp(Left$(t, 4), "Dim ", vbTextCompare) = 0 Then
                    body = Mid$(t, 5)
                Else
                    body = Mid$(t, 8)
                End If
                Dim decls() As String
                decls = Split(body, ",")
                Dim di As Long
                For di = 0 To UBound(decls)
                    Wh2AddVar Owner, out, seenNames, decls(di), "local", sl + li
                Next
            End If
        End If
    Next
    scan("parsedTotal") = out.Count
    If Len(CStr(scan("reason"))) = 0 Then scan("reason") = ""
    Err.Clear
    Set Wh2SourceVars = out
End Function

Public Function Wh2ValueShape(Owner As CodeTools, ByVal expr As String, ByVal frameIndex As Long, ByRef diag As Object) As Object
    Dim rec As Object
    Set rec = CreateObject("Scripting.Dictionary")
    rec("expr") = expr
    On Error Resume Next
    Dim d As Object
    Set d = Wh2EvalRaw(Owner, expr)
    Dim parsed As Boolean
    parsed = False
    If Not d Is Nothing Then
        If d.Exists("valueParsed") Then parsed = CBool(d("valueParsed"))
    End If
    If parsed Then
        rec("value") = d("value")
        Dim tnv As String
        tnv = ""
        If d.Exists("typeName") Then tnv = CStr(d("typeName"))
        rec("typeName") = tnv
        Set rec("inspect") = Wh2Inspect(Owner, expr, frameIndex, diag, tnv)
    Else
        If d Is Nothing Then
            rec("error") = "immediate_eval returned nothing"
        Else
            Dim em As String
            em = ""
            If d.Exists("outputTrimmed") Then em = CStr(d("outputTrimmed"))
            If Len(em) = 0 Then em = "no value parsed (ret=" & CStr(d("ret")) & ")"
            rec("error") = Left$(em, 400)
        End If
    End If
    Err.Clear
    Set Wh2ValueShape = rec
End Function

