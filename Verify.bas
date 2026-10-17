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

Private Function Jq(ByVal k As String) As String
    Jq = Chr$(34) & k & Chr$(34)
End Function

Private Function Js(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Js = Chr$(34) & t & Chr$(34)
End Function
