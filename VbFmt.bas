Attribute VB_Name = "VbFmt"
Option Explicit

' ---------------------------------------------------------------------------
' VB6 source tooling that needs no IDE:
'   * FormatJson  - re-indent a VB file. It ONLY rewrites leading whitespace and
'                   proves it afterwards (the stripped text must be identical),
'                   so it can never corrupt a string, comment or continuation.
'   * RenameJson  - rename an identifier across every component of a .vbp,
'                   skipping string literals, comments and Attribute lines.
' The keyword scanner masks strings/comments first, so keywords that appear
' inside literals never influence indentation or renaming.
' ---------------------------------------------------------------------------

' Re-indents one VB source file. apply=False returns the diff only.
Public Function FormatJson(ByVal path As String, ByVal indentUnit As String, ByVal apply As Boolean) As String
    If Len(path) = 0 Then Err.Raise vbObjectError + 81, "VbFmt", "path is required"
    Dim text As String
    text = ProjShip.ReadTextFile(path)
    If Len(text) = 0 Then Err.Raise vbObjectError + 81, "VbFmt", "file is empty or unreadable: " & path
    If Len(indentUnit) = 0 Then indentUnit = "    "
    Dim eol As String
    eol = LineEndingOf(text)
    Dim lines() As String
    lines = SplitLines(text)
    Dim lo As Long, hi As Long
    lo = LBound(lines)
    hi = UBound(lines)
    Dim out() As String
    ReDim out(lo To hi)
    Dim lvl As Long, i As Long, changed As Long
    lvl = 0
    changed = 0
    For i = lo To hi
        MaybePump
        Dim raw As String, body As String
        raw = lines(i)
        body = LTrim$(raw)
        If Len(body) = 0 Then
            out(i) = ""
        Else
            Dim opens As Long, closes As Long, isLabel As Boolean
            ClassifyLine body, opens, closes, isLabel
            If closes > 0 Then
                If lvl > 0 Then lvl = lvl - 1
            End If
            If isLabel Then
                out(i) = body
            Else
                out(i) = Space$(lvl * Len(indentUnit)) & body
                If opens > 0 Then lvl = lvl + opens
            End If
        End If
        If out(i) <> raw Then changed = changed + 1
    Next
    ' ---- safety proof: formatting may only touch leading whitespace --------
    For i = lo To hi
        MaybePump
        If LTrim$(out(i)) <> LTrim$(lines(i)) Then
            Err.Raise vbObjectError + 82, "VbFmt", "refusing to write: content changed on line " & CStr(i + 1)
        End If
    Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = path
    d("applied") = apply
    d("lines") = hi - lo + 1
    d("changedLines") = changed
    d("sha256Before") = Sha256HexFile(path)
    If apply And changed > 0 Then
        ProjShip.WriteTextFile path, Join(out, eol)
        d("sha256After") = Sha256HexFile(path)
    Else
        d("sha256After") = Sha256HexFile(path)
    End If
    ' TASK-P1-121 item 7 (G-4): state the backup policy instead of leaving it implicit.
    ' format_code writes NO .bak, on purpose - the frozen description says so - so this
    ' is an additive field and the description is deliberately NOT touched.
    d("backup") = False
    d("backupsCreated") = 0
    d("note") = "only leading whitespace is rewritten; the stripped text is verified identical before writing"
    FormatJson = JsonStringify(d)
End Function

' Renames an identifier across all component files of a project. dryRun only
' reports. String literals, comments and Attribute lines are never touched.
Public Function RenameJson(ByVal vbpPath As String, ByVal oldName As String, ByVal newName As String, _
                           ByVal dryRun As Boolean, ByVal skipAttributes As Boolean) As String
    If Len(vbpPath) = 0 Then Err.Raise vbObjectError + 83, "VbFmt", "path is required (the .vbp)"
    If Len(oldName) = 0 Then Err.Raise vbObjectError + 83, "VbFmt", "old_name is required"
    If Len(newName) = 0 Then Err.Raise vbObjectError + 83, "VbFmt", "new_name is required"
    If Not IsIdent(newName) Then Err.Raise vbObjectError + 83, "VbFmt", "new_name is not a valid identifier: " & newName
    If StrComp(oldName, newName, vbTextCompare) = 0 Then Err.Raise vbObjectError + 83, "VbFmt", "old_name and new_name are the same"
    Dim items As Object
    Set items = New Collection
    Dim comps As Object
    Set comps = ProjShip.ComponentPaths(vbpPath)
    Dim total As Long
    total = 0
    Dim attrWarn As Boolean
    attrWarn = False
    Dim v As Variant
    ' TASK-P1-121 item 7 (G-4): the .bak files THIS call creates, named in the receipt.
    Dim baks As Object
    Set baks = New Collection
    For Each v In comps
        Dim f As String
        f = CStr(v)
        Dim txt As String
        txt = ProjShip.ReadTextFile(f)
        If Len(txt) > 0 Then
            Dim before As String
            before = Sha256HexFile(f)
            Dim n As Long
            n = 0
            Dim after As String
            after = ReplaceIdent(txt, oldName, newName, skipAttributes, n, attrWarn)
            If n > 0 Then
                Dim e As Object
                Set e = CreateObject("Scripting.Dictionary")
                e("file") = f
                e("occurrences") = n
                e("sha256Before") = before
                If dryRun Then
                    e("sha256After") = before
                Else
                    ProjShip.WriteTextFile f & ".bak", txt
                    ProjShip.WriteTextFile f, after
                    e("backupFile") = f & ".bak"
                    baks.Add f & ".bak"
                    SeNoteBackup f & ".bak"
                    e("sha256After") = Sha256HexFile(f)
                End If
                items.Add e
                total = total + n
            End If
        End If
    Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("vbp") = vbpPath
    d("oldName") = oldName
    d("newName") = newName
    d("dryRun") = dryRun
    d("totalOccurrences") = total
    d("filesChanged") = items.Count
    Set d("files") = items
    ' TASK-P1-121 item 7 (G-4): enumerate the backups instead of only mentioning them in
    ' prose; the same paths also reach the additive sideEffects block (item 6).
    Set d("backupsCreated") = baks
    If attrWarn Then
        d("warning") = "an Attribute line matches old_name (a module/class name): this tool skips Attribute lines by design - use rename_component for a component rename."
        d("routeHint") = "Component or module names must be renamed with rename_component (it also keeps the .vbp in sync)."
    End If
    d("note") = "Project-wide TEXT rename: string literals, comments and Attribute lines are skipped, but there is no scope analysis, so same-named locals in different procedures are renamed too. Each rewritten file keeps a .bak."
    RenameJson = JsonStringify(d)
End Function

' --------------------------------------------------------------- scanner

' Returns the code part of a line with strings blanked and the comment removed.
Public Function CodePart(ByVal ln As String) As String
    Dim i As Long, inLit As Boolean, ch As String, nxt As String, out As String
    out = ""
    i = 1
    Do While i <= Len(ln)
        MaybePump
        ch = Mid$(ln, i, 1)
        If ch = Chr$(34) Then
            If inLit Then
                inLit = False
            Else
                inLit = True
            End If
            out = out & " "
        ElseIf ch = "'" And Not inLit Then
            Exit Do
        ElseIf inLit Then
            out = out & " "
        Else
            out = out & ch
        End If
        i = i + 1
    Loop
    CodePart = out
End Function

' Replaces whole-word occurrences outside strings/comments. n receives the count.
Private Function ReplaceIdent(ByVal txt As String, ByVal oldName As String, ByVal newName As String, _
                              ByVal skipAttributes As Boolean, ByRef n As Long, ByRef attrWarn As Boolean) As String
    Dim eol As String
    eol = LineEndingOf(txt)
    ' TASK-P1-052: was "out = out & eol / & piece" per line (quadratic on a whole
    ' module). Collect the pieces and Join ONCE. The old code inserted eol only when the
    ' accumulated output was ALREADY non-empty, so leading EMPTY pieces contribute no
    ' separator: find the first non-empty piece and Join from there - that reproduces the
    ' old byte sequence exactly.
    Dim lines() As String, i As Long, j As Long
    lines = SplitLines(txt)
    If UBound(lines) < 0 Then Exit Function
    n = 0
    Dim pieces() As String, np As Long, first As Long, piece As String, ln As String
    ReDim pieces(0 To UBound(lines))
    np = 0
    first = -1
    For i = LBound(lines) To UBound(lines)
        MaybePump
        ln = lines(i)
        If skipAttributes And StartsWith(Trim$(ln), "Attribute ") Then
            If InStr(1, ln, "VB_Name", vbTextCompare) > 0 Then
                If InStr(1, ln, Chr$(34) & oldName & Chr$(34), vbTextCompare) > 0 Then attrWarn = True
            End If
            piece = ln
        Else
            piece = ReplaceOutsideLiterals(ln, oldName, newName, n)
        End If
        pieces(np) = piece
        If first < 0 Then
            If Len(piece) > 0 Then first = np
        End If
        np = np + 1
    Next
    If first < 0 Or np = 0 Then
        ReplaceIdent = ""
        Exit Function
    End If
    Dim keep() As String
    ReDim keep(0 To np - 1 - first)
    For j = first To np - 1
        MaybePump
        keep(j - first) = pieces(j)
    Next
    ReplaceIdent = Join(keep, eol)
End Function

Private Function ReplaceOutsideLiterals(ByVal ln As String, ByVal oldName As String, ByVal newName As String, ByRef n As Long) As String
    Dim i As Long, inLit As Boolean, ch As String, out As String, inComment As Boolean
    Dim L As Long
    L = Len(oldName)
    i = 1
    inLit = False
    inComment = False
    out = ""
    Do While i <= Len(ln)
        MaybePump
        ch = Mid$(ln, i, 1)
        If inComment Then
            out = out & ch
            i = i + 1
        ElseIf inLit Then
            out = out & ch
            If ch = Chr$(34) Then
                If Mid$(ln, i + 1, 1) = Chr$(34) Then
                    out = out & Mid$(ln, i + 1, 1)
                    i = i + 2
                Else
                    inLit = False
                    i = i + 1
                End If
            Else
                i = i + 1
            End If
        ElseIf ch = Chr$(34) Then
            inLit = True
            out = out & ch
            i = i + 1
        ElseIf ch = "'" Then
            inComment = True
            out = out & ch
            i = i + 1
        Else
            If i + L - 1 <= Len(ln) Then
                If StrComp(Mid$(ln, i, L), oldName, vbTextCompare) = 0 Then
                    If IsBoundary(ln, i - 1) And IsBoundary(ln, i + L) Then
                        out = out & newName
                        n = n + 1
                        i = i + L
                    Else
                        out = out & ch
                        i = i + 1
                    End If
                Else
                    out = out & ch
                    i = i + 1
                End If
            Else
                out = out & ch
                i = i + 1
            End If
        End If
    Loop
    ReplaceOutsideLiterals = out
End Function

' A boundary character is anything that cannot be part of an identifier.
Private Function IsBoundary(ByVal ln As String, ByVal pos As Long) As Boolean
    If pos < 1 Then
        IsBoundary = True
        Exit Function
    End If
    If pos > Len(ln) Then
        IsBoundary = True
        Exit Function
    End If
    Dim ch As String
    ch = Mid$(ln, pos, 1)
    IsBoundary = Not IsIdentChar(ch)
End Function

Private Function IsIdentChar(ByVal ch As String) As Boolean
    If ch >= "a" And ch <= "z" Then IsIdentChar = True: Exit Function
    If ch >= "A" And ch <= "Z" Then IsIdentChar = True: Exit Function
    If ch >= "0" And ch <= "9" Then IsIdentChar = True: Exit Function
    If ch = "_" Then IsIdentChar = True
End Function

Private Function IsIdent(ByVal s As String) As Boolean
    IsIdent = False
    If Len(s) = 0 Then Exit Function
    Dim i As Long, ch As String
    ch = Left$(s, 1)
    If Not (IsIdentChar(ch) And Not (ch >= "0" And ch <= "9")) Then Exit Function
    For i = 2 To Len(s)
        MaybePump
        If Not IsIdentChar(Mid$(s, i, 1)) Then Exit Function
    Next
    IsIdent = True
End Function

' --------------------------------------------------------------- indentation

' openN/closes are what the line does to the indent level. Case/Else/ElseIf
' both close and re-open, so they render one level in while leaving the level
' unchanged.
Private Sub ClassifyLine(ByVal body As String, ByRef openN As Long, ByRef closes As Long, ByRef isLabel As Boolean)
    openN = 0
    closes = 0
    isLabel = False
    Dim code As String
    code = Trim$(CodePart(body))
    If Len(code) = 0 Then Exit Sub
    If IsLabelLine(code) Then
        isLabel = True
        Exit Sub
    End If
    Dim u As String
    u = UCase$(code)
    Dim cont As Boolean
    cont = (Right$(Trim$(body), 1) = "_")
    ' conditional compilation
    If Left$(u, 3) = "#IF" Then
        If Left$(u, 4) = "#IF " Then
            openN = 1
        Else
            openN = 1
        End If
        Exit Sub
    End If
    If Left$(u, 6) = "#END I" Then
        closes = 1
        Exit Sub
    End If
    If Left$(u, 6) = "#ELSEI" Or Left$(u, 5) = "#ELSE" Then
        closes = 1
        openN = 1
        Exit Sub
    End If
    ' closers
    If StartsWith(u, "END SUB") Or StartsWith(u, "END FUNCTION") Or StartsWith(u, "END PROPERTY") Then
        closes = 1
        Exit Sub
    End If
    If StartsWith(u, "END IF") Or StartsWith(u, "END SELECT") Or StartsWith(u, "END WITH") Then
        closes = 1
        Exit Sub
    End If
    If StartsWith(u, "END TYPE") Or StartsWith(u, "END ENUM") Then
        closes = 1
        Exit Sub
    End If
    If TokenIs(u, "NEXT") Then
        closes = 1
        Exit Sub
    End If
    If TokenIs(u, "LOOP") Then
        closes = 1
        Exit Sub
    End If
    If TokenIs(u, "WEND") Then
        closes = 1
        Exit Sub
    End If
    ' mid-block lines that render one level in
    If TokenIs(u, "CASE") Then
        closes = 1
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "ELSEIF") Or TokenIs(u, "ELSE") Then
        closes = 1
        openN = 1
        Exit Sub
    End If
    ' openers
    If TokenIs(u, "SELECT CASE") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "WITH") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "DO") Then
        If InStr(u, "LOOP") > 0 Then Exit Sub
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "FOR") Then
        If InStr(u, "NEXT") > 0 Then Exit Sub
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "WHILE") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "TYPE") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "ENUM") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "SUB") Or TokenIs(u, "PUBLIC SUB") Or TokenIs(u, "PRIVATE SUB") Or TokenIs(u, "FRIEND SUB") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "FUNCTION") Or TokenIs(u, "PUBLIC FUNCTION") Or TokenIs(u, "PRIVATE FUNCTION") Or TokenIs(u, "FRIEND FUNCTION") Then
        openN = 1
        Exit Sub
    End If
    If TokenIs(u, "PROPERTY") Or TokenIs(u, "PUBLIC PROPERTY") Or TokenIs(u, "PRIVATE PROPERTY") Or TokenIs(u, "FRIEND PROPERTY") Then
        openN = 1
        Exit Sub
    End If
    ' block If: 'If ... Then' with nothing after Then (or a continuation)
    If TokenIs(u, "IF") Then
        Dim p As Long, rest As String
        p = InStrRev(u, " THEN")
        If p > 0 Then
            rest = Trim$(Mid$(u, p + 5))
            If Len(rest) = 0 Or cont Then openN = 1
        ElseIf cont Then
            openN = 1
        End If
        Exit Sub
    End If
End Sub

Private Function IsLabelLine(ByVal code As String) As Boolean
    IsLabelLine = False
    Dim p As Long
    p = InStr(code, ":")
    If p <= 1 Then Exit Function
    If Len(Trim$(Mid$(code, p + 1))) > 0 Then Exit Function
    Dim head As String, i As Long, ch As String
    head = Trim$(Left$(code, p - 1))
    If Len(head) = 0 Then Exit Function
    If Right$(head, 1) = ")" Then Exit Function
    For i = 1 To Len(head)
        MaybePump
        ch = Mid$(head, i, 1)
        If Not IsIdentChar(ch) Then Exit Function
    Next
    IsLabelLine = True
End Function

Private Function LineEndingOf(ByVal text As String) As String
    If InStr(text, vbCrLf) > 0 Then
        LineEndingOf = vbCrLf
    Else
        LineEndingOf = vbLf
    End If
End Function

Private Function SplitLines(ByVal text As String) As String()
    SplitLines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Function

' Whole-token match: "DO" must not fire on DoEvents, "WITH" not on
' Withdrawal, "NEXT" not on NextValue, "FOR" not on Format.
Private Function TokenIs(ByVal u As String, ByVal token As String) As Boolean
    If StrComp(u, token, vbTextCompare) = 0 Then
        TokenIs = True
        Exit Function
    End If
    If Len(u) > Len(token) Then
        If StrComp(Left$(u, Len(token) + 1), token & " ", vbTextCompare) = 0 Then TokenIs = True
    End If
End Function

Private Function StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    If Len(s) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbTextCompare) = 0)
End Function
