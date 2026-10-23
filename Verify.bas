Attribute VB_Name = "Verify"
Option Explicit

' millisecond clock for the per-segment timing this module now reports.
Private Declare Function GetTickCount Lib "kernel32" () As Long

' ===================== WP-22-04: semantic classification state =====================
' SemanticDiffJson is the only comparison that knows WHICH KIND of difference it found.
' Its answer is published twice: as JSON in the receipt, and in these module variables,
' so VerifyComponent makes its DECISION from the classification itself instead of
' re-parsing the JSON it just built. The add-in is single-threaded and a verify call
' never re-enters itself, so run state here is safe in exactly the way Vb6Lint.bas's
' per-run state already is.
Private m_semAvail As Boolean
Private m_semReal As Long
Private m_semCase As Long
Private m_semWs As Long
Private m_semAttr As Long
Private m_semHeader As Long
Private m_semKind As String

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

    ' ---- WP-22-04: fold the SEMANTIC answer into the DECISION --------------------
    ' ideEqualsDisk stays the BYTE answer and is NOT redefined below. The question a
    ' VB6 caller actually asks is case-insensitive (the compiler compares identifiers
    ' that way), so it gets its own answer - ideEqualsDiskFolded - and only a REAL
    ' difference may fail. foldableOnly means: the disk text could be read AND every
    ' difference folds away (case / trailing whitespace / hidden Attribute lines).
    Dim foldableOnly As Boolean
    foldableOnly = (m_semAvail And (m_semReal = 0))
    Dim diffKind As String
    diffKind = m_semKind
    Dim bres As TBuildResult, buildOk As Long
    buildOk = buildState
    If Len(buildLogPath) > 0 Then
        bres = ParseBuild(buildLogPath, headerOffset)
        If bres.ok Then buildOk = 1 Else buildOk = 0
        buildSource = "log"
    End If
    Dim overall As String
    overall = "PASS"
    ' WP-22-04: a FOLDABLE difference is NOT a proven inconsistency. The measured
    ' defect (MCP-1022-49): ten components reloaded cleanly, verify_scope answered
    ' pass:0 fail:8 with failBasis=ide_disk_diff_proven, and every differing line was
    ' the SAME identifier spelled with different case (m_id / m_Id, Value.Keys /
    ' Value.keys, App.Path / App.path, Mid$ / mid$) - semantically identical to VB6.
    If verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL" Then
        If Not foldableOnly Then overall = "FAIL"
    End If
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
    ' WP-22-04: the classification as TOP-LEVEL fields, so a caller never parses
    ' semanticDiff to branch on it. caseInsensitive is a statement of fact, not a mode:
    ' VB6 identifiers are case-insensitive, so this comparison folds case on BOTH sides.
    sb = sb & Jq("caseInsensitive") & ":true,"
    sb = sb & Jq("diffKind") & ":" & Js(diffKind) & ","
    sb = sb & Jq("caseOnlyLines") & ":" & CStr(m_semCase) & ","
    sb = sb & Jq("whitespaceOnlyLines") & ":" & CStr(m_semWs) & ","
    sb = sb & Jq("attributeOnlyLines") & ":" & CStr(m_semAttr) & ","
    sb = sb & Jq("headerOnlyLines") & ":" & CStr(m_semHeader) & ","
    sb = sb & Jq("realDiffLines") & ":" & CStr(m_semReal) & ","
    sb = sb & Jq("foldableOnly") & ":" & LayB(foldableOnly) & ","
    sb = sb & Jq("ideEqualsDiskFolded") & ":" & LayB(foldableOnly) & ","
    sb = sb & Jq("foldNote") & ":" & Js("foldableOnly=true means every IDE-vs-disk difference is one VB6 itself ignores (identifier case, trailing whitespace, the hidden Attribute lines a pane never shows); it is NOT byte equality - see ideEqualsDisk") & ","
    ' W-D: LAYERED failure facts. Each layer answers on its own; overall/FAIL is derived
    ' from them but a caller can see WHICH layer failed instead of one opaque ok:false.
    Dim lay As String
    lay = "{"
    ' WP-22-04: the layer's failed flag follows the SAME decision as overall, so a
    ' foldable-only difference cannot read as a failed layer while overall says PASS.
    lay = lay & Jq("ideDiskDiff") & ":{" & Jq("failed") & ":" & _
        LayB((verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL") And (Not foldableOnly)) & "," & _
        Jq("ideDisk") & ":" & Js(verdict) & "," & Jq("ideEqualsDisk") & ":" & eqFlag & "," & _
        Jq("foldableOnly") & ":" & LayB(foldableOnly) & "," & Jq("diffKind") & ":" & Js(diffKind) & "," & _
        Jq("detail") & ":" & Js(detail) & "},"
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
    If (verdict = "MISMATCH" Or verdict = "NOT-SYNCED" Or verdict = "READ-FAIL") And (Not foldableOnly) Then
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
    ' WP-22-04: what a lint verdict IS, spliced from the lint engine (ONE source:
    ' Vb6Lint.GateDict). A lint error is a suspicion; the compiler is the authority,
    ' and on a large module a STRUCTURAL error is downgraded there to a warning +
    ' needsHumanReview so it cannot turn this verdict red on its own.
    Dim gateLvl As String
    gateLvl = Vb6Lint.GateLevel
    sb = sb & Jq("gateLevel") & ":" & Js(gateLvl) & ","
    sb = sb & Jq("lintGate") & ":" & Vb6Lint.GateJson() & ","
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
    ' WP-22-04: reset the published classification FIRST, so a run that exits early
    ' (the disk text could not be read) never leaves the PREVIOUS module's answer
    ' behind for a caller to read.
    m_semAvail = False
    m_semReal = 0
    m_semCase = 0
    m_semWs = 0
    m_semAttr = 0
    m_semHeader = 0
    m_semKind = "unavailable"
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
    Dim wsOnly As Long
    Dim headerOnly As Long
    Dim attrOnly As Long
    Dim realN As Long
    Dim diffCount As Long
    Dim extra As Long
    caseOnly = 0
    wsOnly = 0
    headerOnly = 0
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
                ' WP-22-04: an Attribute line skipped on the FIRST row sits in the FILE
                ' HEADER (headerOnly); one skipped later is an inline attribute
                ' (attributeOnly). Both fold, but the contract names both categories.
                If i = 0 Then
                    headerOnly = headerOnly + 1
                Else
                    attrOnly = attrOnly + 1
                End If
                diffCount = diffCount + 1
                If diffCount <= 200 Then
                    If i = 0 Then
                        linesJson = SemAppend(linesJson, i + 1, k + 1, "header-only")
                    Else
                        linesJson = SemAppend(linesJson, i + 1, k + 1, "attribute-only")
                    End If
                End If
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
                ' SemNorm = LCase$(trailing-whitespace-trimmed). Equal here means the
                ' lines differ at most by CASE and/or TRAILING whitespace; the two are
                ' split because they answer different questions for a caller.
                If SemTrimR(ideLine) = SemTrimR(diskLine) Then
                    If ideLine <> diskLine Then
                        wsOnly = wsOnly + 1
                        diffCount = diffCount + 1
                        If diffCount <= 200 Then linesJson = SemAppend(linesJson, i + 1, k + 1, "whitespace-only")
                    End If
                Else
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
    ' WP-22-04: publish the classification (module state + the FOUR named categories).
    m_semAvail = True
    m_semReal = realN
    m_semCase = caseOnly
    m_semWs = wsOnly
    m_semAttr = attrOnly
    m_semHeader = headerOnly
    m_semKind = SemKindOf(diffCount, realN, caseOnly, wsOnly, attrOnly, headerOnly)
    js = "{"
    js = js & SemQ("available") & ":true,"
    js = js & SemQ("semanticEquals") & ":" & SemB(semEq) & ","
    js = js & SemQ("foldableEquals") & ":" & SemB(semEq) & ","
    js = js & SemQ("caseInsensitive") & ":true,"
    js = js & SemQ("diffKind") & ":" & SemS(m_semKind) & ","
    js = js & SemQ("diffCount") & ":" & CStr(diffCount) & ","
    js = js & SemQ("caseOnly") & ":" & CStr(caseOnly) & ","
    js = js & SemQ("attributeOnly") & ":" & CStr(attrOnly) & ","
    js = js & SemQ("whitespaceOnly") & ":" & CStr(wsOnly) & ","
    js = js & SemQ("headerOnly") & ":" & CStr(headerOnly) & ","
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
    SemNorm = LCase$(SemTrimR(ln))
End Function

' Trailing SPACE/TAB only. Leading whitespace and indentation are NOT folded: an
' indentation change is a REAL difference (see the comment above SemanticDiffJson).
Private Function SemTrimR(ByVal ln As String) As String
    Dim t As String
    t = ln
    Do While Len(t) > 0
        If (Right$(t, 1) = " ") Or (Right$(t, 1) = Chr$(9)) Then
            t = Left$(t, Len(t) - 1)
        Else
            Exit Do
        End If
    Loop
    SemTrimR = t
End Function

' The ONE place the classification NAME is decided from the counts.
Private Function SemKindOf(ByVal diffCount As Long, ByVal realN As Long, ByVal caseN As Long, _
        ByVal wsN As Long, ByVal attrN As Long, ByVal headN As Long) As String
    SemKindOf = "none"
    If diffCount <= 0 Then Exit Function
    If realN > 0 Then
        SemKindOf = "real"
        Exit Function
    End If
    If caseN > 0 And wsN = 0 And attrN = 0 And headN = 0 Then
        SemKindOf = "caseOnly"
    ElseIf wsN > 0 And caseN = 0 And attrN = 0 And headN = 0 Then
        SemKindOf = "whitespaceOnly"
    ElseIf attrN > 0 And caseN = 0 And wsN = 0 And headN = 0 Then
        SemKindOf = "attributeOnly"
    ElseIf headN > 0 And caseN = 0 And wsN = 0 And attrN = 0 Then
        SemKindOf = "headerOnly"
    Else
        SemKindOf = "foldableMixed"
    End If
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
