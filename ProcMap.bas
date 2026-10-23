Attribute VB_Name = "ProcMap"
Option Explicit

' ===========================================================================
' COMMENT SEARCH: comment -> procedure mapping helpers (ADDITIVE, flag-gated).
' WHY A SEPARATE MODULE: adding these procedures to CodeTools.cls (a ~1.94 MB
' class) made VB6 6.0 fail with "syntax error" at a line that MOVED with the
' inserted text (evidence: fix/TASK88_COMMENTPROC_*/stub_bisect.py,
' probe_capacity.py). The pure text helpers therefore live in their own module;
' the class only calls them.
' Two sources, never mixed silently:
'   * inside the IDE a component line is mapped by CodeModule.ProcOfLine - the
'     SAME call the IDE's own procedure box uses (ProcMap.AddProcFieldsIde);
'   * for a FILE searched by search_in_files there is no CodeModule, so the map
'     is a LEXICAL scan and the receipt says so (procSource=text_scan).
' A line outside every procedure answers proc="" - it is never guessed.
' ===========================================================================

Public Sub AddProcFieldsIde(ByVal cm As Object, ByVal target As Object, ByVal lineNo As Long, ByVal wantLead As Boolean)
    Dim p As String
    Dim bodyLine As Long
    Dim startLine As Long
    Dim lead As String
    Dim leadN As Long
    p = ""
    bodyLine = 0
    startLine = 0
    On Error Resume Next
    p = cm.ProcOfLine(lineNo, 0)
    Err.Clear
    If Len(p) > 0 Then
        bodyLine = cm.ProcBodyLine(p, 0)
        Err.Clear
        startLine = cm.ProcStartLine(p, 0)
        Err.Clear
    End If
    On Error GoTo 0
    target("proc") = p
    target("procBodyLine") = bodyLine
    target("procStartLine") = startLine
    target("procSource") = "ide.ProcOfLine/ProcBodyLine/ProcStartLine"
    If wantLead Then
        LeadCommentOfIde cm, startLine, bodyLine, lead, leadN
        target("leadingComment") = lead
        target("leadingCommentLines") = leadN
    End If
End Sub

' Leading comment block, DEFINED as: the contiguous run of single-quote comment
' lines immediately ABOVE the procedure header with NO blank line in between (a
' blank line detaches the comment, so it is not leading). ProcStartLine already
' points at the first attached line; the loop re-derives it from the header so
' the rule is visible and testable on its own.
Public Sub LeadCommentOfIde(ByVal cm As Object, ByVal procStartLine As Long, ByVal bodyLine As Long, ByRef outText As String, ByRef outLines As Long)
    outText = ""
    outLines = 0
    If bodyLine <= 1 Then Exit Sub
    Dim i As Long
    Dim s As String
    i = bodyLine - 1
    Do While i >= 1
        s = ""
        On Error Resume Next
        s = cm.Lines(i, 1)
        Err.Clear
        On Error GoTo 0
        If Len(Trim$(s)) = 0 Then Exit Do
        If Left$(LTrim$(s), 1) <> "'" Then Exit Do
        If outLines > 0 Then
            outText = s & vbCrLf & outText
        Else
            outText = s
        End If
        outLines = outLines + 1
        i = i - 1
    Loop
End Sub

' A LEXICAL procedure map for ONE file's lines (owner index per line, 1-based;
' 0 = outside every procedure). LIMIT, declared in the receipt: a header/End
' pair inside a string literal or a comment can confuse the scan, and this is
' NOT the IDE's ProcOfLine.
Public Function ProcLineMapOfText(ByRef lines() As String) As Object
    Dim d As Object
    Dim n As Long
    Dim owner() As Long
    Dim names As Object
    Dim headers As Object
    Dim i As Long
    Dim depth As Long
    Dim cur As Long
    Dim nm As String
    Set d = CreateObject("Scripting.Dictionary")
    d("basis") = "lexical scan of the FILE lines (not the IDE's ProcOfLine)"
    On Error Resume Next
    n = UBound(lines) - LBound(lines) + 1
    If Err.Number <> 0 Then
        Err.Clear
        n = 0
    End If
    On Error GoTo 0
    If n <= 0 Then
        d("owner") = owner
        Set names = New Collection
        Set headers = New Collection
        Set d("names") = names
        Set d("headers") = headers
        Set ProcLineMapOfText = d
        Exit Function
    End If
    ReDim owner(0 To n - 1)
    Set names = New Collection
    Set headers = New Collection
    depth = 0
    cur = 0
    For i = 0 To n - 1
        nm = ""
        If IsProcHeaderLineOf(lines(i), nm) Then
            names.Add nm
            headers.Add i + 1
            cur = names.Count
            depth = 1
            owner(i) = cur
        ElseIf IsProcEndLineOf(lines(i)) Then
            owner(i) = cur
            depth = 0
            cur = 0
        ElseIf depth > 0 Then
            owner(i) = cur
        End If
    Next i
    d("owner") = owner
    Set d("names") = names
    Set d("headers") = headers
    ' Leading-comment ownership: a COMMENT line above a procedure belongs to that procedure
    ' (the whole point of searching by comment). leadStarts(i) = first line of the contiguous
    ' single-quote comment block directly above procedure i's header, or 0 when it has none.
    Dim leadStarts As Object
    Set leadStarts = New Collection
    Dim q As Long, hdrLine As Long, j As Long, keep As Boolean
    For q = 1 To names.Count
        hdrLine = CLng(headers(q))
        j = hdrLine - 1
        keep = True
        Do While j >= 1
            If Len(Trim$(lines(j - 1))) = 0 Then Exit Do
            If Left$(LTrim$(lines(j - 1)), 1) <> "'" Then Exit Do
            j = j - 1
        Loop
        If j <= hdrLine - 1 Then
            leadStarts.Add j + 1
        Else
            leadStarts.Add 0
        End If
    Next q
    Set d("leadStarts") = leadStarts
    Set ProcLineMapOfText = d
End Function

Public Function IsProcHeaderLineOf(ByVal s As String, ByRef outName As String) As Boolean
    Dim t As String
    Dim rest As String
    Dim stripped As Boolean
    Dim k As Long
    Dim ch As String
    IsProcHeaderLineOf = False
    outName = ""
    t = LCase$(Trim$(s))
    If Len(t) = 0 Then Exit Function
    If Left$(t, 1) = "'" Then Exit Function
    Do
        stripped = False
        If Left$(t, 7) = "public " Then
            t = Mid$(t, 8)
            stripped = True
        ElseIf Left$(t, 8) = "private " Then
            t = Mid$(t, 9)
            stripped = True
        ElseIf Left$(t, 7) = "friend " Then
            t = Mid$(t, 8)
            stripped = True
        ElseIf Left$(t, 7) = "global " Then
            t = Mid$(t, 8)
            stripped = True
        ElseIf Left$(t, 7) = "static " Then
            t = Mid$(t, 8)
            stripped = True
        End If
        If Not stripped Then Exit Do
    Loop
    If Left$(t, 4) = "sub " Then
        rest = Trim$(Mid$(t, 5))
    ElseIf Left$(t, 9) = "function " Then
        rest = Trim$(Mid$(t, 10))
    ElseIf Left$(t, 8) = "property" Then
        rest = Trim$(Mid$(t, 9))
        If Left$(rest, 4) = "get " Then
            rest = Trim$(Mid$(rest, 5))
        ElseIf Left$(rest, 4) = "let " Then
            rest = Trim$(Mid$(rest, 5))
        ElseIf Left$(rest, 4) = "set " Then
            rest = Trim$(Mid$(rest, 5))
        Else
            Exit Function
        End If
    Else
        Exit Function
    End If
    outName = ""
    For k = 1 To Len(rest)
        ch = Mid$(rest, k, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            outName = outName & ch
        Else
            Exit For
        End If
    Next k
    If Len(outName) > 0 Then IsProcHeaderLineOf = True
End Function

Public Function IsProcEndLineOf(ByVal s As String) As Boolean
    Dim t As String
    IsProcEndLineOf = False
    t = LCase$(Trim$(s))
    If Left$(t, 7) = "end sub" Then
        IsProcEndLineOf = True
    ElseIf Left$(t, 12) = "end function" Then
        IsProcEndLineOf = True
    ElseIf Left$(t, 12) = "end property" Then
        IsProcEndLineOf = True
    End If
End Function

Public Sub LeadCommentOfLines(ByRef lines() As String, ByVal headerLine As Long, ByRef outText As String, ByRef outLines As Long)
    Dim i As Long
    Dim s As String
    outText = ""
    outLines = 0
    If headerLine <= 1 Then Exit Sub
    i = headerLine - 1
    Do While i >= 1
        s = lines(i - 1)
        If Len(Trim$(s)) = 0 Then Exit Do
        If Left$(LTrim$(s), 1) <> "'" Then Exit Do
        If outLines > 0 Then
            outText = s & vbCrLf & outText
        Else
            outText = s
        End If
        outLines = outLines + 1
        i = i - 1
    Loop
End Sub

Public Sub AddProcFieldsText(ByVal fileMap As Object, ByRef lines() As String, ByVal target As Object, ByVal lineNo As Long, ByVal wantLead As Boolean)
    Dim owner As Variant
    Dim names As Object
    Dim headers As Object
    Dim idx As Long
    Dim pname As String
    Dim hdr As Long
    Dim lead As String
    Dim leadN As Long
    Dim leadStarts As Object
    Dim ls As Long
    Dim probe As Long
    pname = ""
    hdr = 0
    idx = 0
    If fileMap Is Nothing Then
        target("proc") = ""
        target("procSource") = "text_scan_unavailable"
        Exit Sub
    End If
    owner = fileMap("owner")
    Set names = fileMap("names")
    Set headers = fileMap("headers")
    If lineNo >= 1 Then
        On Error Resume Next
        If lineNo <= UBound(owner) + 1 Then idx = CLng(owner(lineNo - 1))
        Err.Clear
        On Error GoTo 0
    End If
    If idx > 0 Then
        If idx <= names.Count Then
            pname = CStr(names(idx))
            hdr = CLng(headers(idx))
        End If
    Else
        ' A hit in a LEADING COMMENT block belongs to the procedure below it -
        ' that is the whole point of searching by comment. Without this a comment
        ' hit would answer proc="" and the feature would map nothing.
        Set leadStarts = fileMap("leadStarts")
        For probe = 1 To names.Count
            ls = 0
            On Error Resume Next
            ls = CLng(leadStarts(probe))
            Err.Clear
            On Error GoTo 0
            If ls > 0 Then
                If lineNo >= ls Then
                    If lineNo < CLng(headers(probe)) Then
                        pname = CStr(names(probe))
                        hdr = CLng(headers(probe))
                        Exit For
                    End If
                End If
            End If
        Next probe
    End If
    target("proc") = pname
    target("procBodyLine") = hdr
    target("procLineSpace") = "file"
    target("procSource") = "text_scan (search_in_files has no IDE CodeModule for an arbitrary file)"
    If wantLead Then
        LeadCommentOfLines lines, hdr, lead, leadN
        target("leadingComment") = lead
        target("leadingCommentLines") = leadN
    End If
End Sub
