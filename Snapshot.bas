Attribute VB_Name = "Snapshot"
Option Explicit

Public Function Sha256OfText(ByVal text As String) As String
    Sha256OfText = Sha256HexTextUtf8(text)
End Function

' the old version compared LINE COUNTS computed by dropping exactly ONE
' trailing empty line per side. VB6 writes the file with an extra trailing blank line
' (measured: disk split_lines=12, pane split_lines=10 for identical content), so a
' freshly saved component answered NOT-SYNCED - a misleading receipt (the very failure
' mode this project keeps hunting). Now the PANE BLOCK is located inside the disk text
' and compared line by line; only trailing blank lines are normalised away, real
' differences are still reported with the differing line. No array is returned from a
' function (VB6 cannot): the line arrays come back ByRef.
Public Function VerifyTextVsDisk(ByVal ideText As String, ByVal diskPath As String, ByVal headerOffset As Long, _
                                ByRef detail As String, ByRef outIdeLines As Long, ByRef outDiskLines As Long, ByRef outOffset As Long, _
                                Optional ByRef outFmtCount As Long = 0, Optional ByRef outFmtFirst As String = "") As String
    Dim b() As Byte, enc As String, diskText As String
    ReadAllBytesTo diskPath, b
    If ByteCount(b) <= 0 Then
        VerifyTextVsDisk = "READ-FAIL"
        detail = "disk empty or unreadable"
        Exit Function
    End If
    enc = DetectEncoding(b)
    DecodeBytes b, enc, diskText
    Dim il() As String, dl() As String
    SplitLinesOf ideText, il
    SplitLinesOf diskText, dl
    TrimTrailingBlank il
    TrimTrailingBlank dl
    Dim ideN As Long, diskN As Long
    ideN = ArrCount(il)
    diskN = ArrCount(dl)
    outIdeLines = ideN
    outDiskLines = diskN
    Dim off As Long
    off = headerOffset
    If off < 0 Then off = 0
    If off > diskN Then off = 0
    outOffset = off
    Dim bad As Long, firstDiff As String, hit As Boolean
    hit = BlockMatches(dl, off, il, ideN, bad, firstDiff)
    If Not hit Then
        ' the caller's header offset may be stale or unresolved: find the pane block by content
        Dim alt As Long
        alt = LocateBlock(dl, il, ideN)
        If alt >= 0 Then
            off = alt
            outOffset = off
            hit = True
            bad = 0
            firstDiff = ""
        End If
    End If
    detail = "disk=" & CStr(diskN) & " ide=" & CStr(ideN) & " headerOffset=" & CStr(off) & " lineSpace=ide (pane and disk are compared as text; trailing blank lines are ignored on both sides)"
    If hit Then
        VerifyTextVsDisk = "OK"
        detail = detail & " ; identical content (encoding " & enc & ")"
        Exit Function
    End If
    ' item 10 (F-c): the SECOND, WHOLE-TEXT judgment, taken after BOTH lookups
    ' failed and BEFORE MISMATCH. VB6 itself drops the empty argument parentheses of a
    ' paren-less `Call X()` when it saves, so a component that had just been saved
    ' answered MISMATCH - a false red (the caller then "fixed" nothing). We answer
    ' MISMATCH-FORMATTING only when N(ide) = N(disk) for the WHOLE text and ideN = diskN,
    ' where N removes exactly those parentheses from a line that IS a paren-less Call,
    ' applied to BOTH sides. No case folding, no whitespace folding, no keyword folding.
    ' ANTI-MASKING: the equality demanded is of the ENTIRE text, so any other byte
    ' difference anywhere in the module keeps the old MISMATCH / NOT-SYNCED answer.
    ' item 10, correction found by the live matrix: the whole-text equality is
    ' of the WHOLE PANE against the WHOLE DISK BODY - i.e. the disk file minus the head the
    ' caller's offset resolved. Comparing against the RAW FILE text would make the verdict
    ' unreachable for every .bas/.cls that carries an Attribute head (ideN can never equal
    ' diskN there), which contradicts the item's own acceptance case. The anti-masking
    ' property is unchanged: the equality demanded is still of the ENTIRE body text.
    ' item 10, second correction from the live matrix. The caller's
    ' headerOffset is 0 whenever it could not be RESOLVED, and it cannot be resolved for
    ' this very case (the block lookup fails precisely because of the Call-parens
    ' difference), so keying the check on `off` made the verdict unreachable.
    ' The natural header offset for a pane/disk pair is bytesN-ideN lines, and N() is
    ' LINE-WISE, so N(disk body from that offset) is just a line-suffix of N(whole disk).
    ' One O(n) normalisation of each side decides it - no scan, no guessing.
    ' ANTI-MASKING is unchanged: the equality demanded is still of the ENTIRE body text,
    ' so any difference other than the paren-less Call spelling keeps MISMATCH.
    Dim nIde As String, nDisk As String
    Dim off2 As Long
    off2 = diskN - ideN
    If off2 >= 0 Then
        nIde = NormCallParens(Join(il, vbLf))
        nDisk = NormCallParens(JoinFrom(dl, off2))
        If nIde = nDisk Then
            off = off2
            outOffset = off
            outFmtCount = FmtDiffCount(il, dl, ideN, off, outFmtFirst)
            VerifyTextVsDisk = "MISMATCH-FORMATTING"
            detail = detail & " ; " & CStr(outFmtCount) & " line(s) differ ONLY by the empty argument parentheses of a paren-less Call statement (VB6 drops them when it saves); the whole text is otherwise byte-identical"
            If Len(outFmtFirst) > 0 Then detail = detail & " ; " & outFmtFirst
            Exit Function
        End If
    End If
    If ideN <= diskN - off Then
        VerifyTextVsDisk = "MISMATCH"
        detail = detail & " ; " & CStr(bad) & " differing line(s); " & firstDiff
    Else
        VerifyTextVsDisk = "NOT-SYNCED"
        detail = detail & " ; the pane has " & CStr(ideN - (diskN - off)) & " line(s) the disk body does not have; " & firstDiff
    End If
End Function

' item 10: N - the ONLY normalisation the new verdict may use. It touches a
' line if and only if the line is a paren-less `Call <name>()`: the empty argument
' parentheses are dropped and nothing else moves (leading indentation is preserved, so an
' indentation difference still breaks the whole-text equality and stays a MISMATCH).
Private Function NormCallParens(ByVal text As String) As String
    Dim ls() As String, out() As String, i As Long
    ls = Split(text, vbLf)
    ReDim out(0 To UBound(ls))
    For i = 0 To UBound(ls)
        MaybePump
        out(i) = NormCallLine(ls(i))
    Next
    NormCallParens = Join(out, vbLf)
End Function

Private Function NormCallLine(ByVal ln As String) As String
    NormCallLine = ln
    Dim t As String
    t = Trim$(ln)
    If Len(t) < 8 Then Exit Function
    If LCase$(Left$(t, 5)) <> "call " Then Exit Function
    If Right$(t, 2) <> "()" Then Exit Function
    Dim core As String
    core = RTrim$(Left$(t, Len(t) - 2))
    If Len(core) <= 5 Then Exit Function
    If InStr(core, "(") > 0 Then Exit Function
    NormCallLine = Left$(ln, Len(ln) - Len(LTrim$(ln))) & core
End Function

Private Function FmtDiffCount(ByRef il() As String, ByRef dl() As String, ByVal n As Long, ByVal off As Long, ByRef outFirst As String) As Long
    Dim i As Long, c As Long
    outFirst = ""
    c = 0
    For i = 0 To n - 1
        MaybePump
        If off + i <= UBound(dl) Then
            If il(i) <> dl(off + i) Then
                c = c + 1
                If Len(outFirst) = 0 Then outFirst = "first formatting diff at line " & CStr(i + 1) & ": ide=[" & il(i) & "] disk=[" & dl(off + i) & "]"
            End If
        End If
    Next
    FmtDiffCount = c
End Function

Private Sub SplitLinesOf(ByVal text As String, ByRef out() As String)
    out = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Sub

Private Sub TrimTrailingBlank(ByRef a() As String)
    Dim n As Long, i As Long
    n = UBound(a)
    Do While n >= 0
        If Len(a(n)) = 0 Then n = n - 1 Else Exit Do
    Loop
    If n < 0 Then
        ReDim a(0 To 0)
        a(0) = ""
        Exit Sub
    End If
    Dim keep() As String
    ReDim keep(0 To n)
    For i = 0 To n
        MaybePump
        keep(i) = a(i)
    Next
    a = keep
End Sub

Private Function ArrCount(ByRef a() As String) As Long
    On Error Resume Next
    ArrCount = 0
    ArrCount = UBound(a) - LBound(a) + 1
End Function

' does il appear in dl starting at off? fills bad/firstDiff even when it returns False
Private Function BlockMatches(ByRef dl() As String, ByVal off As Long, ByRef il() As String, ByVal ideN As Long, ByRef bad As Long, ByRef firstDiff As String) As Boolean
    bad = 0
    firstDiff = ""
    Dim diskN As Long, i As Long, di As Long
    diskN = ArrCount(dl)
    If ideN = 0 Then
        BlockMatches = (diskN = off)
        Exit Function
    End If
    If off < 0 Then Exit Function
    If off >= diskN Then Exit Function
    ' A hidden inline Attribute line exists ONLY in the FILE: a code pane never shows it.
    ' The pane block therefore has to be aligned against the disk body with those lines
    ' SKIPPED, otherwise a correctly saved component reads as MISMATCH. They stay visible
    ' to the caller through the pane->disk map (hiddenInlineAttributeLines).
    di = off
    For i = 0 To ideN - 1
        MaybePump
        Do While di < diskN
            If Left$(Trim$(dl(di)), 10) = "Attribute " Then
                di = di + 1
            Else
                Exit Do
            End If
        Loop
        If di >= diskN Then Exit Function
        If dl(di) <> il(i) Then
            bad = bad + 1
            If Len(firstDiff) = 0 Then firstDiff = "first diff at ide line " & CStr(i + 1) & " (disk line " & CStr(di + 1) & "): ide=[" & il(i) & "] disk=[" & dl(di) & "]"
        End If
        di = di + 1
    Next
    BlockMatches = (bad = 0)
End Function

' find the offset where the pane block appears in the disk lines (-1 = not found)
Private Function LocateBlock(ByRef dl() As String, ByRef il() As String, ByVal ideN As Long) As Long
    LocateBlock = -1
    If ideN = 0 Then Exit Function
    Dim diskN As Long, i As Long, k As Long, b As Long, f As String, key As String
    diskN = ArrCount(dl)
    key = ""
    k = -1
    For i = 0 To ideN - 1
        MaybePump
        If Len(Trim$(il(i))) > 0 Then
            key = il(i)
            k = i
            Exit For
        End If
    Next
    If k < 0 Then Exit Function
    For i = 0 To diskN - ideN
        MaybePump
        If dl(i + k) = key Then
            If BlockMatches(dl, i, il, ideN, b, f) Then
                LocateBlock = i
                Exit Function
            End If
        End If
    Next
End Function

' item 10: join the disk lines from i0 onwards with LF (the common text form
' NormCallParens and the comparison use). VB6 cannot return an array from a function, so
' this is the join helper this module did not yet have.
Private Function JoinFrom(ByRef a() As String, ByVal i0 As Long) As String
    Dim n As Long, parts() As String, i As Long, k As Long
    n = ArrCount(a)
    If i0 < 0 Then i0 = 0
    If i0 >= n Then Exit Function
    ReDim parts(0 To n - 1 - i0)
    k = 0
    For i = i0 To n - 1
        MaybePump
        parts(k) = a(i)
        k = k + 1
    Next
    JoinFrom = Join(parts, vbLf)
End Function

Private Function AsLf(ByVal text As String) As String
    AsLf = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
End Function
