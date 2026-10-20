Attribute VB_Name = "Verify"
Option Explicit

' millisecond clock for the per-segment timing this module now reports.
Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Function VerifyComponent(ByVal compKind As String, ByVal ideText As String, ByVal diskPath As String, ByVal headerOffset As Long, ByVal baselineLines As String, ByVal buildLogPath As String, Optional ByVal compName As String = "", Optional ByVal buildState As Long = -1, Optional ByVal buildSource As String = "", Optional ByVal buildNote As String = "", Optional ByVal extraJson As String = "", Optional ByRef tmLint As Long = 0, Optional ByRef tmDiag As Long = 0, Optional ByRef tmDisk As Long = 0) As String
    ' report where the time of a verify actually goes. The caller
    ' (CodeTools.ToolVerify) splices these into the receipt; 0 means "not measured".
    Dim tA As Long
    Dim errN As Long, warnN As Long, i As Long
    SetBaseline baselineLines
    ' (F2): verify''s receipt carries the FULL diagnostics array (DiagnosticsJson
    ' below), so ask the lint engine for UNLIMITED storage - DiagStored = DiagCount and the
    ' array is byte-identical to before. The two counts now come from the engine''s own
    ' counters rather than a 27,005-iteration scan of the array.
    Vb6Lint.SetStoreCap 0
    tA = GetTickCount()
    Lint ideText, compKind
    errN = Vb6Lint.ErrorCount()
    warnN = Vb6Lint.WarningCount()
    tmLint = GetTickCount() - tA
    Dim lintJson As String
    tA = GetTickCount()
    lintJson = DiagnosticsJson()
    tmDiag = GetTickCount() - tA
    ' the comparison now returns both line counts and the offset it
    ' actually used, and the receipt states the coordinate space explicitly instead
    ' of a bare NOT-SYNCED.
    Dim detail As String, verdict As String
    Dim ideLines As Long, diskLines As Long, usedOffset As Long
    ' item 10 (F-c): the formatting-only diff count/first line, reported as a
    ' structured field so a caller never has to parse the prose in ideDiskDetail.
    Dim fmtCount As Long, fmtFirst As String
    fmtCount = 0
    fmtFirst = ""
    Dim eqFlag As String
    eqFlag = "false"
    If Len(diskPath) > 0 Then
        tA = GetTickCount()
        verdict = VerifyTextVsDisk(ideText, diskPath, headerOffset, detail, ideLines, diskLines, usedOffset, fmtCount, fmtFirst)
        tmDisk = GetTickCount() - tA
    Else
        verdict = "SKIPPED"
        detail = "no disk path"
    End If
    If verdict = "OK" Then eqFlag = "true"
    ' W-D (P4/P11): the SEMANTIC question is separate from the byte question.
    ' ideEqualsDisk stays BYTE-level; semanticDiff says whether the two texts mean the
    ' same thing once keyword case, hidden Attribute lines and trailing whitespace are
    ' normalised away, and classifies every differing line.
    Dim semJson As String
    semJson = SemanticDiffJson(ideText, diskPath, usedOffset)
    Dim bres As TBuildResult, buildOk As Long
    buildOk = buildState
    If Len(buildLogPath) > 0 Then
        bres = ParseBuild(buildLogPath, headerOffset)
        If bres.ok Then buildOk = 1 Else buildOk = 0
        buildSource = "log"
    End If
    Dim overall As String
    overall = "PASS"
    If verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL" Then overall = "FAIL"
    ' item 10 (F-c): MISMATCH-FORMATTING is deliberately NOT added to the FAIL
    ' list above (a formatting-only difference is UNPROVEN, not proven bad) and it is NOT
    ' allowed to inherit PASS either: with a fresh build artifact the ladder below would
    ' otherwise hand out overall=PASS for an ide:=disk= that is demonstrably not equal.
    If verdict = "MISMATCH-FORMATTING" Then
        If overall = "PASS" Then overall = "UNKNOWN"
    End If
    If errN > 0 Then overall = "FAIL"
    If buildOk = 0 Then overall = "FAIL"
    ' PASS now requires POSITIVE build evidence. The build state
    ' used to be left at -1 (unknown) because verify passed buildLogPath="", and
    ' "unknown" still produced overall:PASS - so a module with a hard syntax error
    ' answered PASS / lintErrors:0. lintErrors counts this project's OWN rules, not
    ' the VB6 compiler, so lint agreement alone is not evidence of compilability.
    If overall = "PASS" And buildOk <> 1 Then overall = "UNKNOWN"
    Dim sb As String
    sb = "{"
    ' The first parameter is the component KIND ("bas"/"cls"/"frm"). Emitting it as
    ' "component" is exactly why verify(modLoop) answered component:"bas". Emit the
    ' real NAME as "component" and keep the kind under its own key.
    Dim shownName As String
    shownName = compName
    If Len(shownName) = 0 Then shownName = compKind
    sb = sb & Jq("component") & ":" & Js(shownName) & ","
    sb = sb & Jq("kind") & ":" & Js(compKind) & ","
    sb = sb & Jq("overall") & ":" & Js(overall) & ","
    sb = sb & Jq("ideDisk") & ":" & Js(verdict) & ","
    sb = sb & Jq("ideDiskDetail") & ":" & Js(detail) & ","
    ' the caller must not have to guess. ideEqualsDisk is the boolean
    ' answer; lineSpace/lineNumbersAre say which coordinate space the comparison and
    ' the counts live in; ideLineCount/diskLineCount are the normalised counts.
    sb = sb & Jq("ideEqualsDisk") & ":" & eqFlag & ","
    sb = sb & Jq("lineSpace") & ":" & Js("ide") & ","
    sb = sb & Jq("lineNumbersAre") & ":" & Js("ide") & ","
    sb = sb & Jq("ideLineCount") & ":" & CStr(ideLines) & ","
    sb = sb & Jq("diskLineCount") & ":" & CStr(diskLines) & ","
    sb = sb & Jq("diskBodyOffsetUsed") & ":" & CStr(usedOffset) & ","
    ' item 10 (F-c): additive; ideEqualsDisk stays false for this verdict
    ' because the two texts are NOT equal - only their paren-less Call spelling differs.
    sb = sb & Jq("formattingDiffs") & ":{" & Jq("count") & ":" & CStr(fmtCount) & "," & Jq("first") & ":" & Js(fmtFirst) & "," & Jq("kind") & ":" & Js("paren-less Call <name>() empty argument parentheses") & "},"
    sb = sb & Jq("semanticDiff") & ":" & semJson & ","
    ' W-D: LAYERED failure facts. Each layer answers on its own; overall/FAIL is derived
    ' from them but a caller can see WHICH layer failed instead of one opaque ok:false.
    Dim lay As String
    lay = "{"
    lay = lay & Jq("ideDiskDiff") & ":{" & Jq("failed") & ":" & LayB(verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL") & "," & Jq("ideDisk") & ":" & Js(verdict) & "," & Jq("ideEqualsDisk") & ":" & eqFlag & "," & Jq("detail") & ":" & Js(detail) & "},"
    lay = lay & Jq("lint") & ":{" & Jq("failed") & ":" & LayB(errN > 0) & "," & Jq("lintErrors") & ":" & CStr(errN) & "," & Jq("lintWarnings") & ":" & CStr(warnN) & "," & Jq("baselineHits") & ":" & CStr(BaselineHits()) & "},"
    Dim buildFailed As String
    buildFailed = "false"
    If buildOk = 0 Then buildFailed = "true"
    lay = lay & Jq("build") & ":{" & Jq("failed") & ":" & buildFailed & "," & Jq("buildOk") & ":" & CStr(buildOk) & "," & Jq("buildSource") & ":" & Js(buildSource) & "},"
    lay = lay & Jq("modal") & ":{" & Jq("failed") & ":false," & Jq("note") & ":" & Js("verify never opens a modal dialog") & "}"
    lay = lay & "}"
    sb = sb & Jq("failureLayers") & ":" & lay & ","
    Dim basis As String
    basis = "none"
    If verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL" Then
        basis = "ideDiskDiff"
    ElseIf errN > 0 Then
        basis = "lint"
    ElseIf buildOk = 0 Then
        basis = "build"
    End If
    sb = sb & Jq("overallBasis") & ":" & Js(basis) & ","
    sb = sb & Jq("lintErrors") & ":" & CStr(errN) & ","
    sb = sb & Jq("lintWarnings") & ":" & CStr(warnN) & ","
    sb = sb & Jq("baselineHits") & ":" & CStr(BaselineHits()) & ","
    sb = sb & Jq("buildOk") & ":" & CStr(buildOk) & ","
    sb = sb & Jq("buildSource") & ":" & Js(buildSource) & ","
    sb = sb & Jq("lintNote") & ":" & Js("lintErrors counts this project's own lint rules, NOT the VB6 compiler: a module with a hard syntax error can still report 0 lint errors") & ","
    If Len(buildNote) > 0 Then sb = sb & Jq("buildNote") & ":" & Js(buildNote) & ","
    sb = sb & Jq("diagnostics") & ":" & lintJson
    ' alpha: coords / ideSha256 / artifact fields arrive as raw JSON from
    ' CodeTools so this module does not re-implement the coordinate or artifact logic.
    If Len(extraJson) > 0 Then sb = sb & "," & extraJson
    sb = sb & "}"
    VerifyComponent = sb
End Function

Private Function LayB(ByVal v As Boolean) As String
    LayB = "false"
    If v Then LayB = "true"
End Function

Private Function Jq(ByVal k As String) As String
    Jq = Chr$(34) & k & Chr$(34)
End Function

Private Function Js(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Js = Chr$(34) & t & Chr$(34)
End Function
' =====================================================================
' W-D (P4/P11): SEMANTIC diff - independent of the byte verdict.
' Normalisation: (1) keyword/identifier CASE, (2) hidden Attribute lines
' (VB6 writes "Attribute VB_*" / ".VB_VarHelpID" lines that the pane never shows),
' (3) TRAILING whitespace. Nothing else is folded: an indentation or token change
' is a REAL difference.
' Returns {available, semanticEquals, diffCount, caseOnly, attributeOnly, real,
'          lines:[{paneLine,diskLine,category}], note}.
' =====================================================================

' =====================================================================
' W-D (P4/P11): SEMANTIC diff - independent of the byte verdict.
' Normalisation: (1) CASE, (2) hidden Attribute lines, (3) trailing whitespace.
' Nothing else is folded: indentation or token changes are REAL.
' =====================================================================
Public Function SemanticDiffJson(ByVal ideText As String, ByVal diskPath As String, ByVal usedOffset As Long) As String
    On Error Resume Next
    Dim js As String
    Dim disk As String
    Dim encUsed As String
    Dim whyRead As String
    Dim readOk As Boolean
    disk = ""
    encUsed = ""
    whyRead = ""
    readOk = False
    If Len(diskPath) > 0 Then readOk = Build.ReadSourceText(diskPath, disk, encUsed, whyRead)
    If Len(disk) = 0 Then
        SemanticDiffJson = "{" & SemQ("available") & ":false," & SemQ("reason") & ":" & SemS("the disk text could not be read") & "}"
        Exit Function
    End If
    Dim dl() As String
    Dim il() As String
    dl = SemSplit(disk)
    il = SemSplit(ideText)
    Dim diskN As Long
    Dim ideN As Long
    diskN = UBound(dl) + 1
    ideN = UBound(il) + 1
    Dim doff As Long
    doff = usedOffset
    If doff < 0 Then doff = 0
    If doff > diskN Then doff = 0
    If (doff = 0) And (ideN > 0) And (diskN > 0) Then
        doff = SemFindOffset(dl, il, ideN)
    End If
    Dim caseOnly As Long
    Dim attrOnly As Long
    Dim realN As Long
    Dim diffCount As Long
    Dim extra As Long
    caseOnly = 0
    attrOnly = 0
    realN = 0
    diffCount = 0
    extra = 0
    Dim linesJson As String
    linesJson = ""
    Dim i As Long
    Dim k As Long
    Dim ideLine As String
    Dim diskLine As String
    k = doff
    For i = 0 To ideN - 1
        Do While k <= UBound(dl)
            If SemIsAttr(dl(k)) Then
                attrOnly = attrOnly + 1
                diffCount = diffCount + 1
                If diffCount <= 200 Then linesJson = SemAppend(linesJson, i + 1, k + 1, "attribute-only")
                k = k + 1
            Else
                Exit Do
            End If
        Loop
        If k > UBound(dl) Then
            realN = realN + 1
            diffCount = diffCount + 1
            If diffCount <= 200 Then linesJson = SemAppend(linesJson, i + 1, -1, "real")
        Else
            ideLine = il(i)
            diskLine = dl(k)
            If SemNorm(ideLine) = SemNorm(diskLine) Then
                If ideLine <> diskLine Then
                    caseOnly = caseOnly + 1
                    diffCount = diffCount + 1
                    If diffCount <= 200 Then linesJson = SemAppend(linesJson, i + 1, k + 1, "case-only")
                End If
            Else
                realN = realN + 1
                diffCount = diffCount + 1
                If diffCount <= 200 Then linesJson = SemAppend(linesJson, i + 1, k + 1, "real")
            End If
            k = k + 1
        End If
    Next
    Do While k <= UBound(dl)
        If Not SemIsAttr(dl(k)) Then
            If Len(Trim$(dl(k))) > 0 Then extra = extra + 1
        End If
        k = k + 1
    Loop
    If extra > 0 Then
        realN = realN + extra
        diffCount = diffCount + extra
    End If
    Dim semEq As Boolean
    semEq = False
    If realN = 0 Then semEq = True
    js = "{"
    js = js & SemQ("available") & ":true,"
    js = js & SemQ("semanticEquals") & ":" & SemB(semEq) & ","
    js = js & SemQ("diffCount") & ":" & CStr(diffCount) & ","
    js = js & SemQ("caseOnly") & ":" & CStr(caseOnly) & ","
    js = js & SemQ("attributeOnly") & ":" & CStr(attrOnly) & ","
    js = js & SemQ("real") & ":" & CStr(realN) & ","
    js = js & SemQ("lines") & ":[" & linesJson & "],"
    js = js & SemQ("normalisation") & ":" & SemS("case + hidden Attribute lines + trailing whitespace ONLY")
    js = js & "}"
    SemanticDiffJson = js
    Err.Clear
End Function

Private Function SemQ(ByVal k As String) As String
    SemQ = Chr$(34) & k & Chr$(34)
End Function

Private Function SemS(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    SemS = Chr$(34) & t & Chr$(34)
End Function

Private Function SemB(ByVal v As Boolean) As String
    SemB = "false"
    If v Then SemB = "true"
End Function

Private Function SemNorm(ByVal ln As String) As String
    Dim t As String
    t = ln
    Do While Len(t) > 0
        If (Right$(t, 1) = " ") Or (Right$(t, 1) = Chr$(9)) Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    SemNorm = LCase$(t)
End Function

Private Function SemIsAttr(ByVal ln As String) As Boolean
    Dim t As String
    SemIsAttr = False
    t = LCase$(Trim$(ln))
    If Len(t) = 0 Then Exit Function
    If Left$(t, Len("attribute vb_")) = "attribute vb_" Then SemIsAttr = True
End Function

Private Function SemSplit(ByVal text As String) As String()
    Dim t As String
    Dim arr() As String
    Dim n As Long
    t = Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf)
    arr = Split(t, vbLf)
    n = UBound(arr)
    Do While n >= 0
        If Len(Trim$(arr(n))) = 0 Then
            n = n - 1
        Else
            Exit Do
        End If
    Loop
    If n < 0 Then
        SemSplit = Split("", vbLf)
        Exit Function
    End If
    ReDim Preserve arr(0 To n)
    SemSplit = arr
End Function

Private Function SemAppend(ByVal cur As String, ByVal paneLine As Long, ByVal diskLine As Long, ByVal cat As String) As String
    Dim one As String
    one = "{" & SemQ("paneLine") & ":" & CStr(paneLine) & "," & SemQ("diskLine") & ":" & CStr(diskLine) & "," & SemQ("category") & ":" & SemS(cat) & "}"
    If Len(cur) = 0 Then
        SemAppend = one
    Else
        SemAppend = cur & "," & one
    End If
End Function

Private Function SemFindOffset(ByRef dl() As String, ByRef il() As String, ByVal ideN As Long) As Long
    SemFindOffset = 0
    On Error Resume Next
    Dim diskN As Long
    diskN = UBound(dl) + 1
    Dim j As Long
    Dim i As Long
    Dim ok As Boolean
    If ideN <= 0 Then Exit Function
    For j = 0 To diskN - 1
        If SemIsAttr(dl(j)) Then
            ' skip hidden head lines
        ElseIf SemNorm(dl(j)) = SemNorm(il(0)) Then
            ok = True
            For i = 1 To ideN - 1
                If (j + i) > UBound(dl) Then
                    ok = False
                    Exit For
                End If
                If SemNorm(dl(j + i)) <> SemNorm(il(i)) Then
                    ok = False
                    Exit For
                End If
            Next
            If ok Then
                SemFindOffset = j
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function
