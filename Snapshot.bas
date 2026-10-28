Attribute VB_Name = "Snapshot"
Option Explicit

' v1.16.6: caps for the PROVEN content alignment (TryAlignByContent).
' VB6 requires module-level Const declarations BEFORE the first procedure (measured: a
' Const after an End Function is rejected with 'only comments may appear after End
' Sub/End Function').
Private Const ALIGN_MIN_MATCHES As Long = 2
Private Const ALIGN_MAX_CANDIDATES As Long = 16
Private Const ALIGN_MAX_WORK As Long = 400000

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
' v1.16.6 (HONEST ALIGNMENT): the comparison must PROVE the alignment before it
' may name a differing line. The old flow used the caller's SCALAR headerOffset as
' the only alignment hypothesis; when a REAL difference existed the verbatim block
' lookup failed and the differing-line list was then computed at that guess - a
' fabricated line-by-line verdict (measured: a rewritten disk body reported "first
' diff at ide line 1 (disk line 7)"). Now:
'   1. the caller's offset, then a verbatim content search (as before);
'   2. the whole-text paren-less-Call normalisation (as before);
'   3. NEW: a content-anchor alignment that accepts an offset ONLY when it is the
'      unique best match (>= ALIGN_MIN_MATCHES non-blank pane lines, case-insensitive,
'      blank lines ignored) - then the differing lines are named AT THAT OFFSET;
'   4. otherwise NOT-ALIGNED with aligned:false + the basis, and NO line list.
Public Function VerifyTextVsDisk(ByVal ideText As String, ByVal diskPath As String, ByVal headerOffset As Long, _
                                ByRef detail As String, ByRef outIdeLines As Long, ByRef outDiskLines As Long, ByRef outOffset As Long, _
                                Optional ByRef outFmtCount As Long = 0, Optional ByRef outFmtFirst As String = "", _
                                Optional ByRef outAligned As Boolean = False, Optional ByRef outAlignBasis As String = "", _
                                Optional ByRef outAlignScore As Long = 0) As String
    outAligned = False
    outAlignBasis = ""
    outAlignScore = 0
    Dim b() As Byte, enc As String, diskText As String
    ReadAllBytesTo diskPath, b
    If ByteCount(b) <= 0 Then
        VerifyTextVsDisk = "READ-FAIL"
        detail = "disk empty or unreadable"
        outAlignBasis = "the disk file could not be read, so no comparison was possible"
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
    outAlignBasis = "the caller's headerOffset (" & CStr(off) & ") located the pane block VERBATIM in the disk body"
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
            outAlignBasis = "the pane block was located VERBATIM at disk offset " & CStr(off) & " by content search; the caller's headerOffset was not used"
        End If
    End If
    detail = "disk=" & CStr(diskN) & " ide=" & CStr(ideN) & " headerOffset=" & CStr(off) & " lineSpace=ide (pane and disk are compared as text; trailing blank lines are ignored on both sides)"
    If hit Then
        outAligned = True
        outAlignScore = ideN
        VerifyTextVsDisk = "OK"
        detail = detail & " ; identical content (encoding " & enc & ")"
        Exit Function
    End If
    ' item 10 (F-c): the SECOND, WHOLE-TEXT judgment, taken after BOTH lookups failed
    ' and BEFORE MISMATCH. VB6 itself drops the empty argument parentheses of a
    ' paren-less Call X() when it saves, so a component that had just been saved
    ' answered MISMATCH - a false red. MISMATCH-FORMATTING is answered only when the
    ' ENTIRE pane text equals the ENTIRE disk body after that one normalisation, so
    ' any other byte difference anywhere in the module keeps the old MISMATCH.
    ' The natural header offset is bytesN-ideN lines and N() is LINE-WISE, so
    ' N(disk body from that offset) is just a line-suffix of N(whole disk): one O(n)
    ' normalisation of each side decides it - no scan, no guessing.
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
            outAligned = True
            outAlignBasis = "the whole pane text equals the whole disk body (equal line counts) after the paren-less Call normalisation, which proves the alignment at offset " & CStr(off)
            outAlignScore = ideN
            VerifyTextVsDisk = "MISMATCH-FORMATTING"
            detail = detail & " ; " & CStr(outFmtCount) & " line(s) differ ONLY by the empty argument parentheses of a paren-less Call statement (VB6 drops them when it saves); the whole text is otherwise byte-identical"
            If Len(outFmtFirst) > 0 Then detail = detail & " ; " & outFmtFirst
            Exit Function
        End If
    End If
    ' v1.16.6: PROVE the alignment before naming any differing line.
    Dim aOff As Long, aScore As Long, aBasis As String
    If TryAlignByContent(dl, il, ideN, aOff, aScore, aBasis) Then
        off = aOff
        outOffset = off
        outAligned = True
        outAlignBasis = aBasis
        outAlignScore = aScore
        hit = BlockMatches(dl, off, il, ideN, bad, firstDiff)
        If hit Then
            VerifyTextVsDisk = "OK"
            detail = detail & " ; identical content (encoding " & enc & ")"
            Exit Function
        End If
        If ideN <= diskN - off Then
            VerifyTextVsDisk = "MISMATCH"
            detail = detail & " ; " & aBasis & " ; " & CStr(bad) & " differing line(s); " & firstDiff
        Else
            VerifyTextVsDisk = "NOT-SYNCED"
            detail = detail & " ; " & aBasis & " ; the pane has " & CStr(ideN - (diskN - off)) & " line(s) the disk body does not have; " & firstDiff
        End If
        Exit Function
    End If
    ' No alignment could be PROVEN. Report that instead of a fabricated line list.
    outAligned = False
    outAlignBasis = aBasis
    VerifyTextVsDisk = "NOT-ALIGNED"
    detail = detail & " ; NOT ALIGNED: " & aBasis & ". No differing-line list is produced: naming lines would require an alignment this comparison could not prove."
End Function

' =====================================================================
' v1.16.6: PROVEN content alignment (only used AFTER the verbatim lookups and
' the whole-text formatting normalisation have both failed).
' The candidate offsets come from a CONTENT ANCHOR (the pane's first non-blank
' line, then its last), the score counts case-insensitively equal non-blank pane
' lines, and the offset is accepted only when it is the UNIQUE best score and that
' score is at least ALIGN_MIN_MATCHES. Anything weaker is reported as unproven:
' a tie (which one is "the" alignment?) or a low score cannot justify naming a
' differing line. Bounded by a candidate cap and a work cap so a huge module
' degrades to "unproven" instead of a slow call.
' =====================================================================

Private Function AlignNorm(ByVal ln As String) As String
    AlignNorm = LCase$(Trim$(ln))
End Function

Private Function AlignBlank(ByVal ln As String) As Boolean
    AlignBlank = (Len(Trim$(ln)) = 0)
End Function

Private Function FirstNonBlankIdx(ByRef il() As String, ByVal ideN As Long) As Long
    Dim i As Long
    FirstNonBlankIdx = -1
    For i = 0 To ideN - 1
        MaybePump
        If Not AlignBlank(il(i)) Then
            FirstNonBlankIdx = i
            Exit Function
        End If
    Next
End Function

Private Function LastNonBlankIdx(ByRef il() As String, ByVal ideN As Long) As Long
    Dim i As Long
    LastNonBlankIdx = -1
    For i = ideN - 1 To 0 Step -1
        MaybePump
        If Not AlignBlank(il(i)) Then
            LastNonBlankIdx = i
            Exit Function
        End If
    Next
End Function

Private Function AlignScoreAt(ByRef dl() As String, ByRef il() As String, ByVal ideN As Long, ByVal off As Long, ByRef work As Long) As Long
    Dim i As Long, k As Long, s As Long
    s = 0
    For i = 0 To ideN - 1
        k = off + i
        work = work + 1
        If k >= 0 And k <= UBound(dl) Then
            If AlignNorm(dl(k)) = AlignNorm(il(i)) Then
                If Not AlignBlank(il(i)) Then s = s + 1
            End If
        End If
    Next
    AlignScoreAt = s
End Function

Private Function TryAlignByContent(ByRef dl() As String, ByRef il() As String, ByVal ideN As Long, _
        ByRef outOff As Long, ByRef outScore As Long, ByRef outBasis As String) As Boolean
    TryAlignByContent = False
    outOff = 0
    outScore = 0
    outBasis = ""
    Dim diskN As Long
    diskN = ArrCount(dl)
    If ideN <= 0 Or diskN <= 0 Then
        outBasis = "the component or the disk body is empty, so there is nothing to align"
        Exit Function
    End If
    Dim pass As Long, iA As Long, j As Long, c As Long
    Dim anchor As String
    Dim cand() As Long
    Dim nc As Long
    Dim work As Long, sc As Long, bestScore As Long, secondScore As Long, bestOff As Long, ties As Long
    Dim baseOff As Long
    work = 0
    For pass = 1 To 2
        If pass = 1 Then
            iA = FirstNonBlankIdx(il, ideN)
        Else
            iA = LastNonBlankIdx(il, ideN)
        End If
        If iA >= 0 Then
            anchor = AlignNorm(il(iA))
            nc = 0
            ReDim cand(0 To ALIGN_MAX_CANDIDATES - 1)
            For j = 0 To diskN - 1
                MaybePump
                If AlignNorm(dl(j)) = anchor Then
                    baseOff = j - iA
                    If baseOff >= 0 And baseOff <= diskN - 1 Then
                        If nc < ALIGN_MAX_CANDIDATES Then
                            cand(nc) = baseOff
                            nc = nc + 1
                        End If
                    End If
                End If
            Next
            If nc > 0 Then
                bestScore = 0
                secondScore = 0
                bestOff = 0
                ties = 0
                For c = 0 To nc - 1
                    sc = AlignScoreAt(dl, il, ideN, cand(c), work)
                    If sc > bestScore Then
                        secondScore = bestScore
                        bestScore = sc
                        bestOff = cand(c)
                        ties = 1
                    ElseIf sc = bestScore Then
                        If sc > 0 Then ties = ties + 1
                    ElseIf sc > secondScore Then
                        secondScore = sc
                    End If
                    If work > ALIGN_MAX_WORK Then
                        outBasis = "the content-alignment search reached its work budget (" & CStr(ALIGN_MAX_WORK) & " line comparisons), so NO alignment is claimed"
                        Exit Function
                    End If
                Next
                If bestScore < ALIGN_MIN_MATCHES Then
                    outBasis = "the best content anchor matched only " & CStr(bestScore) & " non-blank pane line(s) (minimum " & CStr(ALIGN_MIN_MATCHES) & "), which does not PROVE an alignment; no differing line is named"
                ElseIf ties > 1 Then
                    outBasis = "the content anchors tied on " & CStr(ties) & " equally good offset(s) (best score " & CStr(bestScore) & "), so choosing between them would be a guess; NO alignment is claimed"
                ElseIf bestScore <= secondScore Then
                    outBasis = "the best content score " & CStr(bestScore) & " is not better than the next best " & CStr(secondScore) & ", so NO alignment is claimed"
                Else
                    outOff = bestOff
                    outScore = bestScore
                    outBasis = "content anchor PROVED the alignment: the unique best offset " & CStr(bestOff) & " matched " & CStr(bestScore) & " non-blank pane line(s) (case-insensitive, blank lines ignored); next best " & CStr(secondScore)
                    TryAlignByContent = True
                    Exit Function
                End If
            Else
                If pass = 2 Then
                    outBasis = "the pane's first/last non-blank line does not occur anywhere in the disk body, so no content anchor exists; NO alignment is claimed"
                End If
            End If
        End If
    Next
    If Len(outBasis) = 0 Then
        outBasis = "no content anchor could be evaluated; NO alignment is claimed"
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
