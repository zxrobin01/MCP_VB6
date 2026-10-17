Attribute VB_Name = "VbAnchor"
Option Explicit

Public Function NormLine(ByVal s As String) As String
    NormLine = LCase$(Trim$(Replace(s, vbTab, " ")))
End Function

Public Function CountLines(ByVal text As String) As Long
    CountLines = UBound(Split(AsLf(text), vbLf)) + 1
End Function

Private Function AsLf(ByVal text As String) As String
    AsLf = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function

Private Function FindUnique(ByRef lines() As String, ByVal oldLine As String, ByRef hit As Long, ByRef cnt As Long)
    Dim i As Long
    cnt = 0
    hit = -1
    For i = 0 To UBound(lines)
        MaybePump
        If NormLine(lines(i)) = NormLine(oldLine) Then
            cnt = cnt + 1
            hit = i
        End If
    Next
End Function

' alpha (edit_by_anchor rejection): expose EVERY hit line so a
' rejection can list the candidates instead of only saying "not unique (N)".
' hits() is 1-based module line numbers; cnt is the number of hits.
Public Function FindHits(ByVal text As String, ByVal oldLine As String, ByRef hits() As Long, ByRef cnt As Long)
    Dim lines() As String
    Dim i As Long
    lines = Split(AsLf(text), vbLf)
    cnt = 0
    ' an EMPTY pane (CountOfLines = 0) is a LEGAL input and simply has
    ' NO hits, but hits() must still come back as a VALID array - VB6's Split("")
    ' returns an EMPTY array (UBound = -1), so "ReDim hits(0 To -1)" raised VB error
    ' 9. This is the root that made edit_by_anchor / apply_edits answer error 9
    ' instead of the structured anchor_not_found an empty component deserves.
    If UBound(lines) < 0 Then
        ReDim hits(0 To 0)
        Exit Function
    End If
    ReDim hits(0 To UBound(lines))
    For i = 0 To UBound(lines)
        MaybePump
        If NormLine(lines(i)) = NormLine(oldLine) Then
            hits(cnt) = i + 1
            cnt = cnt + 1
        End If
    Next
End Function

' ===========================================================================
' one O(N) pass per BATCH instead of one O(N) pass per OP.
'
' The anchor batch path used to call FindHits/FindHitsEx once per op, and each of those
' did "Split(AsLf(text), vbLf)" plus a NormLine (LCase + Trim + Replace) on EVERY line.
' For s ops on an N-line module that is O(s x N) allocations and normalisations: measured
' 337 ms for s=1 but 16,452 ms for s=50 and 33,236 ms for s=500 (which killed the
' connection) on a 27,000-line module.
'
' BuildNormIndex does that work ONCE (raw lines, their normalised forms, and a
' first-hit/count index); FindHitsIndexed then answers an anchor in O(1) when the anchor
' is UNIQUE (the ordinary case). An ambiguous anchor still takes the exact linear scan, so
' the multi-hit list and the occurrence=k semantics stay byte-identical to FindHits.
' ===========================================================================
Public Sub BuildNormIndex(ByVal text As String, ByRef raw() As String, ByRef norm() As String, _
                          ByRef first As Object, ByRef cnts As Object)
    Dim i As Long, n As Long, k As String
    raw = Split(AsLf(text), vbLf)
    ' an EMPTY pane is legal; Split("") is an empty array, so model it as
    ' exactly one blank line (same convention as LoadLines / FindHits).
    If UBound(raw) < 0 Then
        ReDim raw(0 To 0)
        raw(0) = ""
    End If
    n = UBound(raw) + 1
    ReDim norm(0 To n - 1)
    Set first = CreateObject("Scripting.Dictionary")
    Set cnts = CreateObject("Scripting.Dictionary")
    For i = 0 To n - 1
        MaybePump
        k = NormLine(raw(i))
        norm(i) = k
        If first.Exists(k) Then
            cnts(k) = CLng(cnts(k)) + 1
        Else
            first(k) = i + 1
            cnts(k) = 1
        End If
    Next
End Sub

' hits for ONE anchor against an index built by BuildNormIndex.
' cnt = 0 -> hits(0) is a valid (unused) element, exactly like FindHits' empty case.
Public Sub FindHitsIndexed(ByRef norm() As String, ByRef first As Object, ByRef cnts As Object, _
                           ByVal oldLine As String, ByRef hits() As Long, ByRef cnt As Long)
    Dim i As Long, k As String
    k = NormLine(oldLine)
    If Not cnts.Exists(k) Then
        ReDim hits(0 To 0)
        cnt = 0
        Exit Sub
    End If
    cnt = CLng(cnts(k))
    If cnt = 1 Then
        ReDim hits(0 To 0)
        hits(0) = CLng(first(k))
        Exit Sub
    End If
    ' Ambiguous anchor: exact linear fallback so the hit LIST and its order are identical
    ' to FindHits (callers report "anchor not unique" plus every candidate line).
    ReDim hits(0 To UBound(norm))
    cnt = 0
    For i = 0 To UBound(norm)
        If norm(i) = k Then
            hits(cnt) = i + 1
            cnt = cnt + 1
        End If
    Next
End Sub

Public Function AnchorReplaceLine(ByVal text As String, ByVal oldLine As String, ByVal newLine As String, ByRef errMsg As String) As String
    Dim lines() As String, hit As Long, cnt As Long
    lines = Split(AsLf(text), vbLf)
    FindUnique lines, oldLine, hit, cnt
    If cnt = 0 Then
        errMsg = "anchor not found"
        AnchorReplaceLine = ""
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "anchor not unique (" & CStr(cnt) & ")"
        AnchorReplaceLine = ""
        Exit Function
    End If
    lines(hit) = newLine
    AnchorReplaceLine = Join(lines, vbCrLf)
    errMsg = ""
End Function

Public Function AnchorInsertAfter(ByVal text As String, ByVal anchorLine As String, ByVal code As String, ByRef errMsg As String) As String
    Dim lines() As String, ins() As String, out() As String
    Dim hit As Long, cnt As Long, i As Long, j As Long, n As Long
    lines = Split(AsLf(text), vbLf)
    FindUnique lines, anchorLine, hit, cnt
    If cnt = 0 Then
        errMsg = "anchor not found"
        AnchorInsertAfter = ""
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "anchor not unique (" & CStr(cnt) & ")"
        AnchorInsertAfter = ""
        Exit Function
    End If
    ins = Split(AsLf(code), vbLf)
    n = (UBound(lines) + 1) + (UBound(ins) + 1)
    ReDim out(0 To n - 1)
    For i = 0 To hit
        MaybePump
        out(i) = lines(i)
    Next
    For j = 0 To UBound(ins)
        MaybePump
        out(hit + 1 + j) = ins(j)
    Next
    For i = hit + 1 To UBound(lines)
        MaybePump
        out(i + UBound(ins) + 1) = lines(i)
    Next
    AnchorInsertAfter = Join(out, vbCrLf)
    errMsg = ""
End Function

Public Function AnchorDeleteLine(ByVal text As String, ByVal oldLine As String, ByRef errMsg As String) As String
    Dim lines() As String, out() As String
    Dim hit As Long, cnt As Long, i As Long, k As Long
    lines = Split(AsLf(text), vbLf)
    FindUnique lines, oldLine, hit, cnt
    If cnt = 0 Then
        errMsg = "anchor not found"
        AnchorDeleteLine = ""
        Exit Function
    End If
    If cnt > 1 Then
        errMsg = "anchor not unique (" & CStr(cnt) & ")"
        AnchorDeleteLine = ""
        Exit Function
    End If
    If UBound(lines) = 0 Then
        AnchorDeleteLine = ""
        errMsg = ""
        Exit Function
    End If
    ReDim out(0 To UBound(lines) - 1)
    k = 0
    For i = 0 To UBound(lines)
        MaybePump
        If i <> hit Then
            out(k) = lines(i)
            k = k + 1
        End If
    Next
    AnchorDeleteLine = Join(out, vbCrLf)
    errMsg = ""
End Function

' gamma-5/6 (B+C batch): match-aware and scope-aware hit collection.
' matchMode: "normalized" (default: ignores indentation and case, the A batch rule)
'          | "exact" (raw line, binary compare).
' scopeStart/scopeEnd are 1-based IDE lines; 0 means unbounded on that side.
' hits() receives 1-based module line numbers.
Public Function FindHitsEx(ByVal text As String, ByVal oldLine As String, ByVal matchMode As String, _
                           ByVal scopeStart As Long, ByVal scopeEnd As Long, _
                           ByRef hits() As Long, ByRef cnt As Long)
    Dim lines() As String
    Dim i As Long, ln As Long
    lines = Split(AsLf(text), vbLf)
    cnt = 0
    ' same empty-pane root as FindHits (see there): Split("") has
    ' UBound = -1, so "ReDim hits(0 To -1)" was VB error 9.
    If UBound(lines) < 0 Then
        ReDim hits(0 To 0)
        Exit Function
    End If
    ReDim hits(0 To UBound(lines))
    For i = 0 To UBound(lines)
        MaybePump
        ln = i + 1
        If scopeStart > 0 And ln < scopeStart Then GoTo NextLine
        If scopeEnd > 0 And ln > scopeEnd Then GoTo NextLine
        If LineMatches(lines(i), oldLine, matchMode) Then
            hits(cnt) = ln
            cnt = cnt + 1
        End If
NextLine:
    Next
End Function

Public Function LineMatches(ByVal line As String, ByVal want As String, ByVal matchMode As String) As Boolean
    If LCase$(Trim$(matchMode)) = "exact" Then
        LineMatches = (StrComp(line, want, vbBinaryCompare) = 0)
    Else
        LineMatches = (NormLine(line) = NormLine(want))
    End If
End Function

' ===========================================================================
' MULTI-LINE (block) anchors.
' The A batch matched exactly ONE normalized line against the pane, so a caller who
' pasted a two-or-three-line block as the anchor always got anchor_not_found even
' though the block was present verbatim in the code pane. FindBlockHitsEx answers
' the same question for a block: it reports the 1-based START line of every
' position where ALL of the anchor's lines match consecutively, under the same
' match mode and the same scope rules as FindHitsEx, and returns the anchor's line
' count so the caller can know how many lines the match spans.
' ===========================================================================
Public Function AnchorLineCount(ByVal anchorText As String) As Long
    AnchorLineCount = UBound(Split(AsLf(anchorText), vbLf)) + 1
    If AnchorLineCount < 0 Then AnchorLineCount = 0
End Function

Public Function FindBlockHitsEx(ByVal text As String, ByVal anchorText As String, ByVal matchMode As String, _
                                ByVal scopeStart As Long, ByVal scopeEnd As Long, _
                                ByRef hits() As Long, ByRef cnt As Long) As Long
    Dim lines() As String, anc() As String
    Dim i As Long, j As Long, m As Long, ln As Long
    Dim all As Boolean
    lines = Split(AsLf(text), vbLf)
    anc = Split(AsLf(anchorText), vbLf)
    m = UBound(anc) + 1
    cnt = 0
    ReDim hits(0 To 0)
    If m < 1 Then
        FindBlockHitsEx = 0
        Exit Function
    End If
    FindBlockHitsEx = m
    If UBound(lines) < 0 Then Exit Function
    If m > UBound(lines) + 1 Then Exit Function
    ReDim hits(0 To UBound(lines))
    For i = 0 To UBound(lines) - m + 1
        MaybePump
        ln = i + 1
        ' bounds FIRST, elements SECOND (VB6 And does not short-circuit).
        If scopeStart > 0 Then
            If ln < scopeStart Then GoTo NextLine
        End If
        If scopeEnd > 0 Then
            If ln + m - 1 > scopeEnd Then GoTo NextLine
        End If
        all = True
        For j = 0 To m - 1
            If Not LineMatches(lines(i + j), anc(j), matchMode) Then
                all = False
                Exit For
            End If
        Next j
        If all Then
            hits(cnt) = ln
            cnt = cnt + 1
        End If
NextLine:
    Next i
End Function
