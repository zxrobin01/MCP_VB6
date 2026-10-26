Attribute VB_Name = "Build"
Option Explicit

' R-1021-02: run/step disk snapshot budget (declarations section only - VB6
' rejects a module-level Const that appears after the first procedure).
Public Const RUNSNAP_FULL_HASH_MAX As Double = 33554432#
Public Const RUNSNAP_MAX_FILES As Long = 2000

Public Type TBuildResult
    ok As Boolean
    exitCode As Long
    file As String
    reportedLine As Long
    diskLine As Long
    headerOffset As Long
    message As String
    key As String
    raw As String
    ' (i18n): product-first verdict fields
    artifactSize As Double
    okByKeyword As Boolean
    verdict As String
    ' the line the compiler actually wrote (verbatim) and, when nothing could be
    ' recognised, a STRUCTURED unparsed state plus a bounded excerpt of the raw log.  The
    ' receipt must never claim "no error text" - it must say what it could not parse.
    rawLine As String
    unparsed As Boolean
    logExcerpt As String
    ' WHICH rule produced the value - err-marker | fail-marker-up | none
    matchRule As String
    ' the log reader's own facts, so "the file was non-empty but came back empty" can be
    ' settled from the receipt instead of argued (ReadLogAny already reports the encoding).
    logEncoding As String
    logBytes As Double
    ' (step B): the DECODED text the parser actually received, measured - never inferred.
    '   rawLen    = number of CHARACTERS in r.raw.  Read together with logBytes it separates the
    '               states a byte count alone cannot: logBytes>0 AND rawLen=0 proves the file was
    '               read and yielded NOTHING (an unreadable/undecodable log), which is NOT the same
    '               as a genuinely empty file (both = 0, encoding readable).
    '   rawSha256 = sha256 of r.raw re-encoded as UTF-8, so two reads of the same characters agree
    '               and a disagreement is provable without shipping the (possibly huge) text.
    rawLen As Long
    rawSha256 As String
End Type

Private Type TLogClassification
    ' the ONE classification of a compiler log.  Every field is derived in exactly
    ' one place (ClassifyLog) and every consumer reads only these fields.
    markerCount As Long    ' Pass A: how many marker lines the log carries (the counter's n)
    markerLine As String   ' Pass A: the FIRST marker line (what ParseBuild reports)
    failAnchor As Boolean  ' the LAST non-empty line is the failure summary (structural anchor)
    upLine As String       ' Pass B: the line captured above the anchor ("" = nothing captured)
    matchRule As String    ' "none" | "err-marker" | "fail-marker-up"
    rawLine As String      ' markerLine, else upLine - the line a receipt reports
    excerpt As String      ' Pass C: the tail of the last <=10 non-empty lines, when nothing matched
    errorCount As Long     ' THE count: markerCount, else 1 when the anchor holds, else 0
End Type
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ---- path identity ---------------------------------------
' Two spellings can name ONE file: relative/absolute, / vs \, case, doubled
' separators -- and, the case text folding can NEVER see, a junction or symlink
' alias (measured: it is a junction pointing at the reference application). Only
' the filesystem can collapse the alias, so we ask it for the final path of an
' open handle. Desired access is 0 (metadata only) so a file another process is
' holding still opens.
Private Const CF_SHARE_ALL As Long = 7
Private Const CF_OPEN_EXISTING As Long = 3
Private Const CF_BACKUP_SEMANTICS As Long = &H2000000
Private Const CF_INVALID_HANDLE As Long = -1

Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFinalPathNameByHandleW Lib "kernel32" (ByVal hFile As Long, ByVal lpszFilePath As Long, ByVal cchFilePath As Long, ByVal dwFlags As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function RunExe(ByVal exePath As String, ByVal args As String, ByVal timeoutMs As Long, ByRef exitCode As Long) As Boolean
    On Error GoTo fail
    Dim sh As Object, ex As Object, waited As Long
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(Chr$(34) & exePath & Chr$(34) & " " & args)
    waited = 0
    Do While ex.Status = 0
        Sleep 100
        waited = waited + 100
        If waited >= timeoutMs Then
            ex.Terminate
            exitCode = -1
            RunExe = False
            Exit Function
        End If
    Loop
    exitCode = ex.ExitCode
    RunExe = True
    Exit Function
fail:
    exitCode = -1
    RunExe = False
End Function

Public Function ReadLogAnsi(ByVal logPath As String) As String
    ' (F2): the old body ALWAYS decoded as ANSI. A log that is really
    ' UTF-8 (or UTF-16) then came back as U+FFFD replacement characters, so the
    ' compiler message ("syntax error" / the localized equivalent) was unreadable.
    ' ReadLogAny reads it in ITS OWN encoding and reports which one was used.
    Dim s As String, e As String
    ReadLogAny logPath, s, e
    ReadLogAnsi = s
End Function

' (F2): detect the log's real encoding, decode with it, and return the
' name of the encoding that was used so a receipt can PROVE how the text was read.
' DetectEncoding answers "utf-8" only when every byte sequence is valid UTF-8, and a
' GBK (system ANSI) log is not, so a Chinese ANSI log is never mistaken for UTF-8.
Public Sub ReadLogAny(ByVal logPath As String, ByRef out As String, ByRef encUsed As String)
    Dim b() As Byte, n As Long
    out = ""
    encUsed = ""
    On Error Resume Next
    ReadAllBytesTo logPath, b
    n = Encoding.ByteCount(b)
    If Err.Number <> 0 Or n <= 0 Then
        encUsed = "empty_or_unreadable"
        Err.Clear
        Exit Sub
    End If
    ' R-1021-04 (item 4): ONE decode rule, shared with the source-line reader
    ' (Encoding.DecodeUtf8OrAnsi): valid UTF-8 => utf-8, otherwise the system ANSI
    ' codepage (936/ACP). The previous DetectEncoding+DecodeBytes pair here let the
    ' log and the source line be decoded by two different rules.
    Dim decWhy As String
    If Not Encoding.DecodeUtf8OrAnsi(b, out, encUsed, decWhy) Then
        encUsed = "unreadable"
    End If
    Err.Clear
End Sub

' the VERIFIED reader for the WRITE paths (add_code{file=}).  ReadLogAnsi /
' ReadLogAny above stay byte-identical ON PURPOSE: they are the compiler-log reader and
' 's firstError/matchRule evidence is built on them. This one is additive and
' refuses (False + why) instead of ever handing a lossy decode to a code pane.
Public Function ReadSourceText(ByVal path As String, ByRef out As String, ByRef encUsed As String, ByRef why As String) As Boolean
    Dim b() As Byte, n As Long
    out = "": encUsed = "": why = ""
    ReadSourceText = False
    On Error Resume Next
    ReadAllBytesTo path, b
    n = Encoding.ByteCount(b)
    If Err.Number <> 0 Then
        why = "input_file_unreadable": Err.Clear: Exit Function
    End If
    Err.Clear
    If n <= 0 Then
        ' an empty/unreadable file keeps its OLD meaning: nothing to append, no refusal
        encUsed = "empty_or_unreadable"
        ReadSourceText = True
        Exit Function
    End If
    If Not Encoding.DecodeVerified(b, encUsed, out, why) Then Exit Function
    ReadSourceText = True
End Function

Public Function HeaderOffset(ByVal diskLines As Long, ByVal ideLines As Long) As Long
    HeaderOffset = diskLines - ideLines
End Function

Private Function CErr() As String
    CErr = ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H9519) & ChrW$(&H8BEF)
End Function
Private Function COk() As String
    COk = ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H6210) & ChrW$(&H529F)
End Function
Private Function CNLine() As String
    CNLine = ChrW$(&H884C)
End Function
Private Function CSyn() As String
    CSyn = ChrW$(&H8BED) & ChrW$(&H6CD5) & ChrW$(&H9519) & ChrW$(&H8BEF)
End Function
Private Function CUndef() As String
    CUndef = ChrW$(&H53D8) & ChrW$(&H91CF) & ChrW$(&H672A) & ChrW$(&H5B9A) & ChrW$(&H4E49)
End Function
' the failure SUMMARY line the compiler writes last ('<exe>' ...).
Private Function CFail() As String
    ' ROOT CAUSE (measured): the full stop here was written as ChrW(&HA1A3) - the GBK BYTE
    ' value of the full-width period - but the decoded log text carries the UNICODE character
    ' U+3002 (。).  The five-character marker could therefore never match, so Pass B never fired and
    ' the receipt fell back to log-unparsed.  The other constants (CErr/COk/CFail2) use real Unicode
    ' code points and were always fine; this one was the byte/character confusion.
    CFail = ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H5931) & ChrW$(&H8D25) & ChrW$(&H3002)
End Function
' 生成失败 - the OTHER localized FAILURE SUMMARY line /make can write (the first one is 编译失败。).
' this is a failure MARKER only, never a sentence to recognise an error by.
Private Function CFail2() As String
    ' 生成失败 - the OTHER localized failure summary /make writes
    CFail2 = ChrW$(&H751F) & ChrW$(&H6210) & ChrW$(&H5931) & ChrW$(&H8D25)
End Function
' ---- ONE shared line recognizer ------------------------------------------
' Both ParseBuild and CountBuildErrors call THIS, so the count and the first error can never
' diverge.  Localized markers are now case-INsensitive and the English ones are the same set.
Public Function IsCompileErrorLine(ByVal ln As String) As Boolean
    ' MARKERS ONLY (structure, not sentences).  The project-level message '必须有启动窗体 或
    ' Sub Main()' is deliberately NOT a marker here: it may only ever be the VERBATIM text that
    ' Pass B captures, so this parser can never be tuned to one project's wording.
    If InStr(1, ln, CErr(), vbTextCompare) > 0 Then IsCompileErrorLine = True
    If InStr(1, ln, "compile error", vbTextCompare) > 0 Then IsCompileErrorLine = True
End Function
' The fail summary is a STRUCTURAL ANCHOR: /make always ENDS with it, so the caller asks this
' predicate ONLY about the LAST non-empty line (never a whole-file scan for "failed", which would
' mistake an error sentence containing that word for the summary).
Public Function IsCompileFailLine(ByVal ln As String) As Boolean
    If InStr(1, ln, CFail(), vbTextCompare) > 0 Then IsCompileFailLine = True
    If InStr(1, ln, CFail2(), vbTextCompare) > 0 Then IsCompileFailLine = True
    If InStr(1, ln, "failed", vbTextCompare) > 0 Then IsCompileFailLine = True
    If InStr(1, ln, "build failed", vbTextCompare) > 0 Then IsCompileFailLine = True
End Function
' A banner / success / bare-path / bare-time line is never an error line even when it sits right
' above the failure summary, so Pass B must not promote it to firstError.
Private Function IsBannerLine(ByVal ln As String) As Boolean
    Dim s As String
    s = Trim$(ln)
    If Len(s) = 0 Then IsBannerLine = True: Exit Function
    If InStr(1, s, "Microsoft (R) Visual Basic", vbTextCompare) > 0 Then IsBannerLine = True: Exit Function
    If InStr(1, s, "Copyright (C)", vbTextCompare) > 0 Then IsBannerLine = True: Exit Function
    If InStr(1, s, CErr(), vbTextCompare) > 0 Then Exit Function
    If IsCompileFailLine(s) Then IsBannerLine = True: Exit Function
    If Len(s) <= 3 Then IsBannerLine = True: Exit Function
    If Mid$(s, 2, 1) = ":" Then IsBannerLine = True: Exit Function
End Function

Public Function ErrorKey(ByVal msg As String) As String
    ' (i18n): the localized words are kept, but English keywords
    ' are accepted too, so a VB6 in another language still classifies the error.
    If InStr(msg, CSyn()) > 0 Or InStr(1, msg, "syntax error", vbTextCompare) > 0 Then
        ErrorKey = "syntax"
    ElseIf InStr(msg, CUndef()) > 0 Or InStr(1, msg, "variable not defined", vbTextCompare) > 0 Then
        ErrorKey = "undefined-variable"
    ElseIf Len(Trim$(msg)) = 0 Then
        ErrorKey = "unknown"
    Else
        ErrorKey = "other"
    End If
End Function

Public Function ParseBuild(ByVal logPath As String, ByVal headerOffset As Long) As TBuildResult
    Dim r As TBuildResult
    r.ok = False
    r.exitCode = -999
    r.headerOffset = headerOffset
    ' take the encoding the reader ACTUALLY used (ReadLogAnsi discarded it), and the
    ' real byte size on disk - both are receipt facts, not inferences.
    Dim encUsed As String
    ReadLogAny logPath, r.raw, encUsed
    r.logEncoding = encUsed
    ' (step B): publish the reader's OWN output.  These two fields are the reason a
    ' "the log is non-empty but came back empty" claim can be settled from the receipt alone:
    ' logBytes is the SIZE ON DISK, rawLen is the number of characters the reader produced.
    r.rawLen = Len(r.raw)
    r.rawSha256 = Sha256.Sha256HexTextUtf8(r.raw)
    On Error Resume Next
    r.logBytes = FileLen(logPath)
    On Error GoTo 0
    Dim t As String
    t = r.raw
    r.okByKeyword = (InStr(t, COk()) > 0)
    If (Not r.okByKeyword) And Len(Trim$(t)) > 0 Then
        If InStr(1, t, "success", vbTextCompare) > 0 Then r.okByKeyword = True
    End If
    r.ok = r.okByKeyword
    ' every recognition rule now runs in ONE place (ClassifyLog).  This function no
    ' longer owns a second copy of the marker scan or of the fail-summary anchor - it only turns
    ' the ONE classification into the receipt fields.
    Dim c As TLogClassification
    ClassifyLog t, c, True
    r.matchRule = c.matchRule
    If Len(c.rawLine) > 0 Then
        r.rawLine = c.rawLine
        If c.matchRule = "err-marker" Then
            ' a marker line: the message is what FOLLOWS the marker, the line itself is the fallback
            r.message = MsgAfter(c.rawLine)
            If Len(Trim$(r.message)) = 0 Then r.message = Trim$(c.rawLine)
            FileLineFrom c.rawLine, r.file, r.reportedLine
        Else
            ' a Pass B capture IS the message, verbatim - it is not a marker sentence
            r.message = Trim$(c.rawLine)
            FileLineFrom r.message, r.file, r.reportedLine
        End If
        r.key = ErrorKey(r.message)
        If headerOffset >= 0 And r.reportedLine > 0 Then r.diskLine = r.reportedLine + headerOffset + 1
    End If
    ' Pass C: still nothing recognised -> record a STRUCTURED unparsed state and a
    ' bounded excerpt of the raw log.  An empty message must never reach the caller, because the
    ' caller used to turn that emptiness into a false statement ('no error text').
    If Len(r.rawLine) = 0 And Len(r.message) = 0 And (Not r.ok) Then
        r.unparsed = True
        r.logExcerpt = c.excerpt
    End If
    ParseBuild = r
End Function

' =====================================================================
' ONE log classifier - the ONLY place a compiler log is classified.
'   ParseBuild and CountBuildErrors used to spell the SAME rules twice: split the log,
'   scan it for a marker line (Pass A), and - when nothing matched - treat the LAST
'   non-empty line as the structural failure summary and step up to the error text
'   (Pass B).  Two copies of one rule can fork, and they already had: the Lead measured
'   a log whose error is reachable ONLY through Pass B being counted as 0 errors while
' firstError had a value. The fix removed the MEASURED fork; this
'   makes the fork structurally impossible - a rule changed here moves the parsed receipt
'   AND the count in the same edit, and no consumer owns a second marker table.
'   The message pump is NOT a classification rule: ParseBuild runs inside the IDE and pumps
'   between lines, CountBuildErrors never pumped.  It is passed in as a flag, so the pump
'   stays with the caller and is not duplicated either.
' =====================================================================
Private Sub ClassifyLog(ByVal raw As String, ByRef c As TLogClassification, ByVal pump As Boolean)
    Dim lines() As String, i As Long, last As Long, ln As String
    Dim j As Long, steps As Long
    c.markerCount = 0
    c.markerLine = ""
    c.failAnchor = False
    c.upLine = ""
    c.matchRule = "none"
    c.rawLine = ""
    c.excerpt = ""
    c.errorCount = 0
    lines = Split(Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    ' ---- ONE pass over the log: the marker rule (Pass A) and the anchor input --------
    ' Split("") returns an array whose UBound is -1, so this loop simply does not run for a
    ' 0-byte log; every walk below is bounded the same way (: VB6's And does NOT
    ' short-circuit, so a bound that guards an index must stand on its own line).
    last = -1
    For i = 0 To UBound(lines)
        ' the pump fires exactly where the old Pass A fired it: on every line up to and
        ' including the first marker line (which is where ParseBuild used to Exit For).
        If pump Then
            If Len(c.markerLine) = 0 Then MaybePump
        End If
        ln = lines(i)
        If Len(Trim$(ln)) > 0 Then last = i
        If IsCompileErrorLine(ln) Then
            c.markerCount = c.markerCount + 1
            If Len(c.markerLine) = 0 Then c.markerLine = ln
        End If
    Next i
    ' ---- Pass B: no marker line -> the failure summary is a STRUCTURAL ANCHOR -------
    ' /make ENDS with its summary line, so the anchor is the LAST non-empty line - never
    ' "the first line anywhere that contains failed".  From there we may step up at most 3
    ' non-empty lines, skipping banners, to find the error text.
    If Len(c.markerLine) = 0 Then
        If last >= 0 Then
            If IsCompileFailLine(lines(last)) Then
                c.failAnchor = True
                j = last - 1
                steps = 0
                Do While j >= 0 And steps < 3
                    If Len(Trim$(lines(j))) > 0 Then
                        steps = steps + 1
                        If Not IsBannerLine(lines(j)) Then Exit Do
                    End If
                    j = j - 1
                Loop
                If j >= 0 Then
                    If steps > 0 Then
                        If Not IsBannerLine(lines(j)) Then
                            c.upLine = lines(j)
                            c.matchRule = "fail-marker-up"
                        End If
                    End If
                End If
            End If
        End If
    End If
    ' ---- the ONE published verdict ---------------------------------------------------
    If Len(c.markerLine) > 0 Then
        c.rawLine = c.markerLine
        c.matchRule = "err-marker"
    ElseIf Len(c.upLine) > 0 Then
        c.rawLine = c.upLine
    End If
    ' The count is DEFINED here, from the same facts that produced rawLine: marker lines
    ' when there are any, else exactly ONE when the structural anchor holds (a log that
    ' reaches Pass B carries one error - the line Pass B captures).
    If c.markerCount > 0 Then
        c.errorCount = c.markerCount
    ElseIf c.failAnchor Then
        c.errorCount = 1
    End If
    ' ---- Pass C: nothing recognised -> a bounded tail excerpt ------------------------
    If Len(c.rawLine) = 0 Then
        Dim ex As String, seen As Long
        seen = 0
        If UBound(lines) >= 0 Then
            For i = UBound(lines) To 0 Step -1
                If Len(Trim$(lines(i))) > 0 Then
                    ex = lines(i) & vbCrLf & ex
                    seen = seen + 1
                    If seen >= 10 Then Exit For
                End If
            Next
        End If
        c.excerpt = ex
    End If
End Sub

' the file/line STRUCTURE rule, in one place.  A quoted component name AND a line
' number are what may become file/reportedLine; a quoted token inside an error sentence
' (e.g. the argument of a localized undefined-variable message) must NOT become a file name,
' and a project-level error must leave both ABSENT rather than filling reportedLine = 0 as if
' it were a real line.  Both Pass A and Pass B used to spell this rule separately.
Private Sub FileLineFrom(ByVal ln As String, ByRef outFile As String, ByRef outLine As Long)
    If Len(BetweenQuotes(ln)) > 0 And LineAfter(ln) > 0 Then
        outFile = BetweenQuotes(ln)
        outLine = LineAfter(ln)
    End If
End Sub

' how many compile errors does a VB6 /out log carry?
' One build stops at the FIRST error, but a caller still needs a count to judge
' whether the log is a single failure or a cascade.: the count is now read
' straight off the SAME classification ParseBuild uses, so the two can no longer fork
' (the fork: a Chinese log whose error only Pass B could reach counted 0 while
' firstError had a value).  No pump here - this function never pumped.
Public Function CountBuildErrors(ByVal raw As String) As Long
    If Len(raw) = 0 Then Exit Function
    Dim c As TLogClassification
    ClassifyLog raw, c, False
    CountBuildErrors = c.errorCount
End Function


' the compiler's line numbers are CODE-PANE lines. When the
' component is not open in the IDE there is no live header offset and diskLine came
' back 0 - which is exactly the user's case: they built a project the IDE did not
' hold and the log said "line 26955". The header is therefore counted from the file
' on disk: the leading VERSION / BEGIN..END / Attribute block VB6 hides from the pane.
Public Function HeaderOffsetFromFile(ByVal path As String) As Long
    HeaderOffsetFromFile = -1
    On Error Resume Next
    If Len(path) = 0 Then Exit Function
    If Dir$(path) = "" Then Exit Function
    Dim fso As Object, ts As Object, t As String, n As Long, inBegin As Boolean
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(path, 1)
    n = 0
    inBegin = False
    Do While Not ts.AtEndOfStream
        t = Trim$(ts.ReadLine)
        If Len(t) = 0 Then Exit Do
        If inBegin Then
            n = n + 1
            If LCase$(t) = "end" Then inBegin = False
        ElseIf InStr(1, t, "Attribute ", vbTextCompare) = 1 Then
            n = n + 1
        ElseIf InStr(1, t, "VERSION ", vbTextCompare) = 1 Then
            n = n + 1
        ElseIf LCase$(t) = "begin" Then
            n = n + 1
            inBegin = True
        ElseIf LCase$(t) = "end" Then
            n = n + 1
        Else
            Exit Do
        End If
        If n > 200 Then Exit Do
    Loop
    ts.Close
    HeaderOffsetFromFile = n
End Function

Private Function BetweenQuotes(ByVal s As String) As String
    Dim a As Long, b As Long
    a = InStr(s, "'")
    If a = 0 Then a = InStr(s, Chr$(34))
    If a = 0 Then Exit Function
    b = InStr(a + 1, s, "'")
    If b = 0 Then b = InStr(a + 1, s, Chr$(34))
    If b = 0 Then Exit Function
    BetweenQuotes = Mid$(s, a + 1, b - a - 1)
End Function

Private Function LineAfter(ByVal s As String) As Long
    Dim p As Long, i As Long, ch As String, num As String, off As Long
    p = InStr(s, CNLine() & " ")
    If p = 0 Then
        p = InStr(1, s, "line ", vbTextCompare)
        If p = 0 Then Exit Function
        off = 5
    Else
        off = 2
    End If
    For i = p + off To Len(s)
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            num = num & ch
        Else
            If Len(num) > 0 Then Exit For
        End If
    Next
    LineAfter = Val(num)
End Function

Private Function MsgAfter(ByVal s As String) As String
    Dim p As Long, off As Long
    p = InStr(s, " : ")
    If p > 0 Then
        off = 3
    Else
        p = InStr(s, ":")
        If p = 0 Then Exit Function
        off = 1
    End If
    MsgAfter = Trim$(Mid$(s, p + off))
End Function

Public Function DiskLineOf(ByVal ideLine As Long, ByVal headerOffset As Long) As Long
    DiskLineOf = ideLine + headerOffset
End Function

Public Function CoordsJson(ByVal ideLine As Long, ByVal headerOffset As Long) As String
    CoordsJson = "{" & Chr$(34) & "ideLine" & Chr$(34) & ":" & CStr(ideLine) & "," & Chr$(34) & "diskLine" & Chr$(34) & ":" & CStr(DiskLineOf(ideLine, headerOffset)) & "," & Chr$(34) & "headerOffset" & Chr$(34) & ":" & CStr(headerOffset) & "}"
End Function

' ---------------------------------------------------------------------------
' (i18n): product-first build verdict.
'
' Success used to be decided by InStr(log, "编译成功") - a Chinese-only word. On a
' VB6 in any other language every build looked failed. The compiled ARTIFACT is
' language-independent, so it decides: the output file must exist, be non-empty,
' and be no older than the build. The localized/English keyword is retained only
' as an auxiliary signal (and as the sole signal when no artifact path is known,
' e.g. a project with an empty ExeName32).
' ---------------------------------------------------------------------------
Public Function ArtifactPathOf(ByVal vbpPath As String) As String
    On Error Resume Next
    Dim fso As Object, ts As Object, ln As String, nm As String
    If Len(vbpPath) = 0 Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(vbpPath, 1)
    Do While Not ts.AtEndOfStream
        ln = ts.ReadLine
        If InStr(1, ln, "ExeName32=", vbTextCompare) = 1 Then
            nm = Replace(Mid$(ln, Len("ExeName32=") + 1), Chr$(34), "")
            Exit Do
        End If
    Loop
    ts.Close
    ' (found while building the junction fixture): TextStream.ReadLine
    ' leaves a stray CR behind when a hand-edited .vbp uses CR CR LF line ends. The
    ' artifact path then ended in CR, Dir$ could not see it, and a successful in-IDE
    ' Make was reported as artifact-missing/ok:false. Strip control characters first.
    nm = Replace(Replace(Replace(nm, vbCr, ""), vbLf, ""), vbTab, "")
    nm = Trim$(nm)
    If Len(nm) = 0 Then Exit Function
    If InStr(nm, ":") > 0 Then
        ArtifactPathOf = nm
    Else
        ArtifactPathOf = fso.GetParentFolderName(vbpPath) & Chr$(92) & nm
    End If
End Function

Public Function ArtifactBuilt(ByVal artifactPath As String, ByVal timeBefore As Double, ByRef sizeNow As Double) As Boolean
    On Error Resume Next
    sizeNow = 0
    If Len(artifactPath) = 0 Then Exit Function
    If Dir$(artifactPath) = "" Then Exit Function
    sizeNow = FileLen(artifactPath)
    If sizeNow <= 0 Then Exit Function
    If timeBefore > 0 Then
        ' FileDateTime has 1-second resolution: allow 2s of slack so a very fast
        ' build never looks stale
        If FileDateTime(artifactPath) < (timeBefore - 2 / 86400#) Then Exit Function
    End If
    ArtifactBuilt = True
End Function

' ---------------------------------------------------------------------------
' incremental build feedback.
' Fingerprint = sha256(.vbp bytes) + sorted sha256 of every component file the
' .vbp references, plus /d defines. Disk based and language independent.
' ---------------------------------------------------------------------------
Public Function FileSha256(ByVal path As String) As String
    ' (F1): delegate; the WHY-carrying form below is what lets a caller
    ' distinguish "no digest" from "the digest of an empty input".
    Dim why As String
    why = ""
    FileSha256 = FileSha256Why(path, why)
End Function

' (F1): a sha256 that can NEVER silently be the digest of an EMPTY input.
' Measured defect: an EXCLUSIVE-locked artifact was reported as
' sha256=E3B0C44298FC1C149AFBF4C8996FB92427AE41E4649B934CA495991B7852B855
' (the well-known digest of the empty string) while output.exists=true and size=16384,
' because ReadAllBytesTo swallows the Open failure and leaves the byte array empty.
' A wrong VALUE is far more dangerous than a missing one, so here an unreadable or
' short-read file answers "" plus a reason instead.
Public Function FileSha256Why(ByVal path As String, ByRef why As String) As String
    why = ""
    FileSha256Why = ""
    If Len(path) = 0 Then
        why = "no_path_given"
        Exit Function
    End If
    If Dir$(path) = "" Then
        why = "file_not_found"
        Exit Function
    End If
    Dim fnum As Integer, n As Long
    Dim b() As Byte
    On Error GoTo fail
    fnum = FreeFile
    Open path For Binary Access Read As #fnum
    n = LOF(fnum)
    If n <= 0 Then
        Close #fnum
        FileSha256Why = Sha256.Sha256HexBytes(b)
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Get #fnum, , b
    Close #fnum
    If Encoding.ByteCount(b) <> n Then
        why = "short_read: the file reports " & CStr(n) & " byte(s) but " & CStr(Encoding.ByteCount(b)) & " were read; no digest is reported instead of the empty-input digest"
        Exit Function
    End If
    FileSha256Why = Sha256.Sha256HexBytes(b)
    Exit Function
fail:
    On Error Resume Next
    Close #fnum
    why = "read_failed: the file could not be opened for reading (another process holds it exclusively, or access was denied) - no digest is reported instead of the digest of an empty input"
End Function

Public Function VbpComponentFiles(ByVal vbpPath As String) As Object
    On Error Resume Next
    Dim res As Object
    Set res = New Collection
    If Len(vbpPath) = 0 Then Set VbpComponentFiles = res: Exit Function
    If Dir$(vbpPath) = "" Then Set VbpComponentFiles = res: Exit Function
    Dim fso As Object, ts As Object, ln As String, kw As String, rhs As String
    Dim p As Long, q As Long, nm As String, rel As String
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(vbpPath, 1)
    Do While Not ts.AtEndOfStream
        ln = Trim$(ts.ReadLine)
        p = InStr(ln, "=")
        If p > 1 Then
            kw = LCase$(Left$(ln, p - 1))
            If kw = "module" Or kw = "class" Or kw = "form" Or kw = "mdiform" _
               Or kw = "usercontrol" Or kw = "propertypage" Or kw = "userdocument" Or kw = "designer" Then
                rhs = Mid$(ln, p + 1)
                q = InStr(rhs, ";")
                If q > 0 Then
                    nm = Trim$(Left$(rhs, q - 1))
                    rel = Trim$(Mid$(rhs, q + 1))
                    If Len(nm) > 0 And Len(rel) > 0 Then
                        Dim d As Object
                        Set d = CreateObject("Scripting.Dictionary")
                        d("name") = nm
                        If InStr(rel, ":") > 0 Then
                            d("path") = rel
                        Else
                            d("path") = fso.GetParentFolderName(vbpPath) & Chr$(92) & rel
                        End If
                        res.Add d
                    End If
                End If
            End If
        End If
    Loop
    ts.Close
    Set VbpComponentFiles = res
End Function

' Text-only normalisation: unify separators, collapse runs of them, and keep a
' LEADING pair of backslashes (the UNC marker) and an extended prefix intact.
Public Function FoldPathText(ByVal p As String) As String
    Dim i As Long, ch As String, t As String, prevSep As Boolean, startI As Long
    If Len(p) = 0 Then FoldPathText = "": Exit Function
    t = ""
    startI = 1
    prevSep = False
    If Left$(p, 2) = "\\" Then
        t = "\\"
        prevSep = True
        startI = 3
    End If
    For i = startI To Len(p)
        MaybePump
        ch = Mid$(p, i, 1)
        If ch = "/" Then ch = "\"
        If ch = "\" Then
            If Not prevSep Then t = t & "\"
            prevSep = True
        Else
            t = t & ch
            prevSep = False
        End If
    Next
    FoldPathText = t
End Function

' Final on-disk spelling of an EXISTING file or directory. This is the only way
' to see through a junction/symlink alias. Returns "" when the path cannot be
' opened (missing, illegal, or refused) -- callers then fall back to text folding.
Public Function FinalPathOf(ByVal p As String) As String
    On Error Resume Next
    FinalPathOf = ""
    If Len(p) = 0 Then Exit Function
    Dim h As Long, n As Long, buf As String
    h = CreateFileW(StrPtr(p), 0&, CF_SHARE_ALL, 0&, CF_OPEN_EXISTING, 0&, 0&)
    ' a junction is a DIRECTORY: that open needs FILE_FLAG_BACKUP_SEMANTICS
    If h = CF_INVALID_HANDLE Then h = CreateFileW(StrPtr(p), 0&, CF_SHARE_ALL, 0&, CF_OPEN_EXISTING, CF_BACKUP_SEMANTICS, 0&)
    If Err.Number <> 0 Then Err.Clear: Exit Function
    If h = CF_INVALID_HANDLE Then Exit Function
    buf = String$(1024, vbNullChar)
    n = GetFinalPathNameByHandleW(h, StrPtr(buf), 1024, 0&)
    CloseHandle h
    If Err.Number <> 0 Then Err.Clear: Exit Function
    If n <= 0 Or n > 1024 Then Exit Function
    buf = Left$(buf, n)
    ' strip the extended prefix the API always prepends; the caller needs a plain
    ' path (VB6.EXE /make and Dir$ do not understand the \\?\ form)
    If Left$(buf, 8) = "\\?\UNC\" Then
        buf = "\\" & Mid$(buf, 9)
    ElseIf Left$(buf, 4) = "\\?\" Then
        buf = Mid$(buf, 5)
    End If
    FinalPathOf = buf
End Function

' Canonical IDENTITY of a path: the filesystem's final spelling when the target
' exists, otherwise the text-folded spelling. Compare the result with
' StrComp(a, b, vbTextCompare) = 0 -- never with a bare =, and never by LCase$
' alone (that is what missed the junction alias this card is about).
Public Function CanonicalPath(ByVal p As String) As String
    On Error Resume Next
    CanonicalPath = ""
    If Len(p) = 0 Then Exit Function
    Dim f As String
    f = FinalPathOf(p)
    If Len(f) > 0 Then
        CanonicalPath = f
    Else
        CanonicalPath = FoldPathText(p)
    End If
End Function

Public Function BuildFingerprint(ByVal vbpPath As String, ByVal defines As String, ByRef compsOut As Object) As Object
    On Error Resume Next
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Dim comps As Object
    Set comps = CreateObject("Scripting.Dictionary")
    r("vbpHash") = FileSha256(vbpPath)
    Dim files As Object
    Set files = VbpComponentFiles(vbpPath)
    Dim i As Long, d As Object, cat As String
    Dim lines() As String
    If files.Count > 0 Then
        ReDim lines(0 To files.Count - 1)
        For i = 1 To files.Count
            MaybePump
            If Util.PumpAbortDue() Then Exit Function
            Set d = files.Item(i)
            Dim sh As String
            sh = FileSha256(d("path"))
            comps(d("name")) = sh
            lines(i - 1) = LCase$(d("name")) & "=" & sh
        Next
        ' plain insertion sort: the component count is small
        Dim a As Long, b2 As Long, tmp As String
        For a = LBound(lines) To UBound(lines) - 1
            MaybePump
            If Util.PumpAbortDue() Then Exit Function
            For b2 = a + 1 To UBound(lines)
                MaybePump
                If Util.PumpAbortDue() Then Exit Function
                If lines(b2) < lines(a) Then
                    tmp = lines(a): lines(a) = lines(b2): lines(b2) = tmp
                End If
            Next
        Next
        For i = LBound(lines) To UBound(lines)
            MaybePump
            If Util.PumpAbortDue() Then Exit Function
            cat = cat & "|" & lines(i)
        Next
    End If
    r("hash") = Sha256.Sha256HexTextUtf8(r("vbpHash") & "|" & defines & cat)
    ' the cache key must be the project IDENTITY, not its spelling.
    r("key") = CanonicalPath(vbpPath) & "|" & LCase$(defines)
    Set compsOut = comps
    Set BuildFingerprint = r
End Function

Private Function BuildCachePath() As String
    BuildCachePath = TempDir() & Chr$(92) & "vb6mcp_build_cache.txt"
End Function

' Loads the last build record into c (a Dictionary). Returns False when absent.
Public Function CacheLoad(ByRef c As Object) As Boolean
    On Error Resume Next
    CacheLoad = False
    Dim p As String
    p = BuildCachePath()
    If Dir$(p) = "" Then Exit Function
    Dim fso As Object, ts As Object, ln As String, sp As Long, kk As String, vv As String
    Dim comps As Object
    Set comps = CreateObject("Scripting.Dictionary")
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.OpenTextFile(p, 1)
    Do While Not ts.AtEndOfStream
        ln = ts.ReadLine
        If Left$(ln, 5) = "COMP " Then
            vv = Mid$(ln, 6)
            sp = InStr(vv, " ")
            If sp > 0 Then comps(Mid$(vv, 1, sp - 1)) = Mid$(vv, sp + 1)
        Else
            sp = InStr(ln, " ")
            If sp > 1 Then
                kk = Left$(ln, sp - 1)
                vv = Mid$(ln, sp + 1)
                Select Case kk
                    Case "KEY": c("key") = vv
                    Case "HASH": c("hash") = vv
                    Case "OK": c("ok") = (vv = "1")
                    Case "EC": c("exitCode") = Val(vv)
                    Case "ART": c("artifact") = vv
                    Case "SIZE": c("size") = Val(vv)
                    Case "LOG": c("logPath") = vv
                    Case "AT": c("at") = Val(vv)
                    Case "VERDICT": c("verdict") = vv
                    Case "OKKW": c("okByKeyword") = (vv = "1")
                    ' FIX-24 (F6): which source the build compiled, and whether pane and disk
                    ' disagreed. Optional: an older cache file simply does not carry them.
                    Case "CF": c("compiledFrom") = vv
                    Case "PDM": c("paneDiskMismatch") = (vv = "1")
                    Case "PDMC": c("paneDiskMismatchCount") = Val(vv)
                End Select
            End If
        End If
    Loop
    ts.Close
    Set c("comps") = comps
    ' FIX-24 (F6): defaults so a cache record written BEFORE this field existed reads as
    ' "unknown compiled source" instead of raising.
    If Not c.Exists("compiledFrom") Then c("compiledFrom") = ""
    If Not c.Exists("paneDiskMismatch") Then c("paneDiskMismatch") = False
    If Not c.Exists("paneDiskMismatchCount") Then c("paneDiskMismatchCount") = 0
    CacheLoad = True
End Function

Public Sub CacheSave(ByVal c As Object)
    On Error Resume Next
    Dim fso As Object, ts As Object, k As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(BuildCachePath(), True)
    ts.WriteLine "KEY " & CStr(c("key"))
    ts.WriteLine "HASH " & CStr(c("hash"))
    If c("ok") Then ts.WriteLine "OK 1" Else ts.WriteLine "OK 0"
    ts.WriteLine "EC " & CStr(c("exitCode"))
    ts.WriteLine "ART " & CStr(c("artifact"))
    ts.WriteLine "SIZE " & CStr(c("size"))
    ts.WriteLine "LOG " & CStr(c("logPath"))
    ts.WriteLine "AT " & CStr(c("at"))
    ts.WriteLine "VERDICT " & CStr(c("verdict"))
    If c("okByKeyword") Then ts.WriteLine "OKKW 1" Else ts.WriteLine "OKKW 0"
    ' FIX-24 (F6): persist the compiled-source axis (additive lines; older readers ignore them).
    ts.WriteLine "CF " & CStr(c("compiledFrom"))
    If c("paneDiskMismatch") Then ts.WriteLine "PDM 1" Else ts.WriteLine "PDM 0"
    ts.WriteLine "PDMC " & CStr(c("paneDiskMismatchCount"))
    If c.Exists("comps") Then
        Dim comps As Object
        Set comps = c("comps")
        For Each k In comps.Keys
            ts.WriteLine "COMP " & CStr(k) & " " & CStr(comps(k))
        Next
    End If
    ts.Close
End Sub

' name -> change diff against the previous build's component hashes
Public Function ChangedComponents(ByVal comps As Object, ByVal cache As Object) As Object
    On Error Resume Next
    Dim res As Object
    Set res = New Collection
    Set ChangedComponents = res
    If Not cache.Exists("comps") Then Exit Function
    Dim prior As Object
    Set prior = cache("comps")
    If prior Is Nothing Then Exit Function
    Dim k As Variant, d As Object
    For Each k In comps.Keys
        If Not prior.Exists(k) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("name") = CStr(k): d("change") = "added"
            res.Add d
        ElseIf CStr(prior(k)) <> CStr(comps(k)) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("name") = CStr(k): d("change") = "changed"
            res.Add d
        End If
    Next
    For Each k In prior.Keys
        If Not comps.Exists(k) Then
            Set d = CreateObject("Scripting.Dictionary")
            d("name") = CStr(k): d("change") = "removed"
            res.Add d
        End If
    Next
End Function

Public Function NamesOf(ByVal coll As Object) As Object
    Dim res As Object
    Set res = New Collection
    Dim i As Long
    If Not coll Is Nothing Then
        Dim it As Object
        For i = 1 To coll.Count
            Set it = coll.Item(i)
            res.Add CStr(it("name"))
        Next
    End If
    Set NamesOf = res
End Function

Public Sub ApplyArtifactVerdict(ByRef r As TBuildResult, ByVal artifactPath As String, ByVal timeBefore As Double)
    Dim sz As Double
    sz = 0
    r.artifactSize = 0
    If ArtifactBuilt(artifactPath, timeBefore, sz) Then
        r.artifactSize = sz
        r.ok = True
        r.verdict = "artifact"
        Exit Sub
    End If
    r.artifactSize = sz
    If Len(artifactPath) = 0 Then
        ' nothing to check on disk: the keyword signal is all we have
        r.verdict = "log-keyword-only"
        r.ok = r.okByKeyword
    Else
        r.verdict = "artifact-missing"
        r.ok = False
    End If
End Sub
' =====================================================================
' R-1021-02: DISK snapshot helpers for run/step observability.
' DISK-ONLY (no IDE objects): the file set comes from the .vbp text, and every
' digest comes from FileSha256Why, whose "" means "no digest" (never the digest
' of an empty input). Budget: a project whose files total <= RUNSNAP_FULL_HASH_MAX
' is hashed COMPLETELY on both sides ("full"); above it the mode degrades to
' "gated": the .vbp and the dirty components are hashed, and a file is hashed
' AFTER the call only when its size/mtime moved. A digest that was not taken is
' reported as "" plus a reason - never invented.
' =====================================================================

Public Function DiskProjectFiles(ByVal vbpPath As String) As Object
    Set DiskProjectFiles = DiskProjectFilesEx(vbpPath, Nothing)
End Function

' R-1021-02: the file set MUST cover every component file. The .vbp text alone is
' not enough: a line like "Form=FormLive.frm" carries NO "name; file" pair, so a
' text-only parser drops that component (measured on the live fixture: 3 of 4
' files were covered). ideFiles (from VBComponents.FileNames) is therefore the
' PRIMARY source; the .vbp text only adds files the IDE did not report.
Public Function DiskProjectFilesEx(ByVal vbpPath As String, ByVal ideFiles As Object) As Object
    Dim res As Object
    Dim seen As Object
    Dim d As Object
    Dim c As Variant
    Dim cp As String
    Dim low As String
    Dim sib As String
    Set res = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    If Len(vbpPath) > 0 Then
        If Dir$(vbpPath) <> "" Then
            Set d = CreateObject("Scripting.Dictionary")
            d("name") = "(project)"
            d("path") = vbpPath
            d("kind") = "vbp"
            res.Add d
            seen(LCase$(vbpPath)) = True
        End If
    End If
    If Not ideFiles Is Nothing Then
        For Each c In ideFiles
            cp = CStr(c("path"))
            If Len(cp) > 0 Then
                If Not seen.Exists(LCase$(cp)) Then
                    Set d = CreateObject("Scripting.Dictionary")
                    d("name") = CStr(c("name"))
                    d("path") = cp
                    d("kind") = "component"
                    res.Add d
                    seen(LCase$(cp)) = True
                End If
            End If
            MaybePump
            If Util.PumpAbortDue() Then Exit Function
        Next
    End If
    If Len(vbpPath) > 0 Then
        If Dir$(vbpPath) <> "" Then
            Dim comps As Object
            Set comps = VbpComponentFiles(vbpPath)
            For Each c In comps
                cp = CStr(c("path"))
                If Len(cp) > 0 Then
                    If Not seen.Exists(LCase$(cp)) Then
                        Set d = CreateObject("Scripting.Dictionary")
                        d("name") = CStr(c("name"))
                        d("path") = cp
                        d("kind") = "component"
                        res.Add d
                        seen(LCase$(cp)) = True
                    End If
                End If
                MaybePump
                If Util.PumpAbortDue() Then Exit Function
            Next
        End If
    End If
    ' paired binary resources: the resource file is a real component file too
    Dim i As Long
    For i = 1 To res.Count
        Set d = res.Item(i)
        If CStr(d("kind")) = "component" Then
            cp = CStr(d("path"))
            low = LCase$(cp)
            sib = ""
            If Right$(low, 4) = ".frm" Then sib = Left$(cp, Len(cp) - 4) & ".frx"
            If Right$(low, 4) = ".ctl" Then sib = Left$(cp, Len(cp) - 4) & ".frx"
            If Right$(low, 4) = ".pag" Then sib = Left$(cp, Len(cp) - 4) & ".pgx"
            If Len(sib) > 0 Then
                If Dir$(sib) <> "" Then
                    If Not seen.Exists(LCase$(sib)) Then
                        Set d = CreateObject("Scripting.Dictionary")
                        d("name") = "(resource) " & sib
                        d("path") = sib
                        d("kind") = "resource"
                        res.Add d
                        seen(LCase$(sib)) = True
                    End If
                End If
            End If
        End If
    Next
    Do While res.Count > RUNSNAP_MAX_FILES
        res.Remove res.Count
    Loop
    Err.Clear
    Set DiskProjectFilesEx = res
End Function

Public Function DiskFileEntry(ByVal path As String, ByVal withSha As Boolean) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = path
    d("exists") = False
    d("bytes") = 0
    d("mtime") = 0
    d("mtimeText") = ""
    d("sha") = ""
    d("shaWhy") = ""
    d("hashed") = False
    On Error Resume Next
    If Len(path) = 0 Then
        d("shaWhy") = "no_path"
        Set DiskFileEntry = d
        Exit Function
    End If
    If Dir$(path) = "" Then
        d("shaWhy") = "file_not_found"
        Set DiskFileEntry = d
        Exit Function
    End If
    d("exists") = True
    d("bytes") = FileLen(path)
    d("mtime") = FileDateTime(path)
    d("mtimeText") = Format$(FileDateTime(path), "yyyy-mm-dd hh:nn:ss")
    If withSha Then
        Dim why As String
        why = ""
        d("sha") = FileSha256Why(path, why)
        d("shaWhy") = why
        d("hashed") = (Len(CStr(d("sha"))) > 0)
    Else
        d("shaWhy") = "not_hashed(budget)"
    End If
    Err.Clear
    Set DiskFileEntry = d
End Function

Public Function DiskSnapTake(ByVal files As Object, ByVal shaPaths As Object, ByRef meta As Object) As Object
    Dim snap As Object
    Set snap = New Collection
    Dim t0 As Long
    t0 = Util.TickMs()
    Dim total As Double
    total = 0
    Dim f As Variant
    Dim sz As Double
    Dim mode As String
    Dim mustHash As Boolean
    On Error Resume Next
    For Each f In files
        sz = 0
        If Dir$(CStr(f("path"))) <> "" Then sz = FileLen(CStr(f("path")))
        If Err.Number <> 0 Then Err.Clear
        total = total + sz
        MaybePump
        If Util.PumpAbortDue() Then Exit Function
    Next
    meta("hashMode") = "gated"
    If total <= RUNSNAP_FULL_HASH_MAX Then meta("hashMode") = "full"
    meta("totalBytes") = total
    meta("fileCount") = files.Count
    meta("budgetBytes") = RUNSNAP_FULL_HASH_MAX
    If total > RUNSNAP_FULL_HASH_MAX Then
        meta("gatedReason") = "the project's files total " & CStr(CLng(total)) & " B, above the " & CStr(CLng(RUNSNAP_FULL_HASH_MAX)) & " B full-hash budget"
    End If
    mode = CStr(meta("hashMode"))
    For Each f In files
        mustHash = (mode = "full")
        If Not mustHash Then
            If CStr(f("kind")) = "vbp" Then mustHash = True
            If Not shaPaths Is Nothing Then
                If shaPaths.Exists(LCase$(CStr(f("path")))) Then mustHash = True
            End If
        End If
        snap.Add DiskFileEntry(CStr(f("path")), mustHash)
        MaybePump
        If Util.PumpAbortDue() Then Exit Function
    Next
    meta("snapshotMs") = Util.TickMs() - t0
    meta("checkedFiles") = snap.Count
    Err.Clear
    Set DiskSnapTake = snap
End Function

' In "gated" mode the AFTER snapshot hashes a file only when its size/mtime
' moved against the BEFORE snapshot; this fills exactly those digests.
Public Sub DiskSnapFillChanged(ByVal prevSnap As Object, ByVal curSnap As Object)
    On Error Resume Next
    Dim prevBy As Object
    Set prevBy = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim a As Object
    Dim b As Object
    Dim k As String
    Dim why As String
    For i = 1 To prevSnap.Count
        Set a = prevSnap.Item(i)
        prevBy(LCase$(CStr(a("path")))) = i
    Next
    For i = 1 To curSnap.Count
        Set b = curSnap.Item(i)
        If Not CBool(b("hashed")) Then
            k = LCase$(CStr(b("path")))
            If prevBy.Exists(k) Then
                Set a = prevSnap.Item(CLng(prevBy(k)))
                If CDbl(a("bytes")) <> CDbl(b("bytes")) Or CDbl(a("mtime")) <> CDbl(b("mtime")) Then
                    why = ""
                    b("sha") = FileSha256Why(CStr(b("path")), why)
                    b("shaWhy") = why
                    b("hashed") = (Len(CStr(b("sha"))) > 0)
                End If
            End If
        End If
        MaybePump
        If Util.PumpAbortDue() Then Exit Sub
    Next
    Err.Clear
End Sub

' Content changes -> writes[]; a digest-identical rewrite -> touched[].
Public Function DiskSnapDiff(ByVal prevSnap As Object, ByVal curSnap As Object, ByVal attribution As String) As Object
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    Dim writes As Object
    Dim touched As Object
    Set writes = New Collection
    Set touched = New Collection
    On Error Resume Next
    Dim prevBy As Object
    Set prevBy = CreateObject("Scripting.Dictionary")
    Dim i As Long
    Dim a As Object
    Dim b As Object
    Dim before As Object
    Dim k As String
    Dim changed As Boolean
    Dim w As Object
    Dim t As Object
    For i = 1 To prevSnap.Count
        Set a = prevSnap.Item(i)
        prevBy(LCase$(CStr(a("path")))) = i
    Next
    For i = 1 To curSnap.Count
        Set b = curSnap.Item(i)
        k = LCase$(CStr(b("path")))
        Set before = Nothing
        If prevBy.Exists(k) Then Set before = prevSnap.Item(CLng(prevBy(k)))
        changed = False
        If before Is Nothing Then
            changed = CBool(b("exists"))
        Else
            If CBool(before("exists")) <> CBool(b("exists")) Then
                changed = True
            ElseIf CBool(b("exists")) Then
                If Len(CStr(before("sha"))) > 0 And Len(CStr(b("sha"))) > 0 Then
                    changed = (StrComp(CStr(before("sha")), CStr(b("sha")), vbTextCompare) <> 0)
                Else
                    If CDbl(before("bytes")) <> CDbl(b("bytes")) Then changed = True
                    If CDbl(before("mtime")) <> CDbl(b("mtime")) Then changed = True
                End If
            End If
        End If
        If changed Then
            Set w = CreateObject("Scripting.Dictionary")
            w("file") = CStr(b("path"))
            If before Is Nothing Then
                w("beforeBytes") = 0
                w("beforeSha") = ""
                w("beforeShaNote") = "the file did not exist before this call"
            Else
                w("beforeBytes") = before("bytes")
                w("beforeSha") = before("sha")
                If Len(CStr(before("sha"))) = 0 Then w("beforeShaNote") = CStr(before("shaWhy"))
            End If
            w("afterBytes") = b("bytes")
            w("afterSha") = b("sha")
            If Len(CStr(b("sha"))) = 0 Then w("afterShaNote") = CStr(b("shaWhy"))
            w("reason") = "pre-run-save"
            w("attribution") = attribution
            writes.Add w
        ElseIf (Not before Is Nothing) And CBool(b("exists")) Then
            If CDbl(before("mtime")) <> CDbl(b("mtime")) Or CDbl(before("bytes")) <> CDbl(b("bytes")) Then
                If Len(CStr(before("sha"))) > 0 And Len(CStr(b("sha"))) > 0 Then
                    If StrComp(CStr(before("sha")), CStr(b("sha")), vbTextCompare) = 0 Then
                        Set t = CreateObject("Scripting.Dictionary")
                        t("file") = CStr(b("path"))
                        t("beforeBytes") = before("bytes")
                        t("afterBytes") = b("bytes")
                        t("mtimeBefore") = before("mtimeText")
                        t("mtimeAfter") = b("mtimeText")
                        t("shaSame") = True
                        touched.Add t
                    End If
                End If
            End If
        End If
        MaybePump
        If Util.PumpAbortDue() Then Exit Function
    Next
    Set out("writes") = writes
    Set out("touched") = touched
    out("checkedFiles") = curSnap.Count
    Err.Clear
    Set DiskSnapDiff = out
End Function

' =====================================================================
' WP-22-04 (MCP-1022-50): WHERE the artifact identity is sampled.
'
' Defect (measured): make_project(route=ide) answered output{size:5972021,
' sha256:"EF46BB4B..."} while an external Get-FileHash of the SAME file returned
' "EAC18630..." with an identical size - i.e. output{} described the artifact as it looked
' right AFTER THE COMPILE and BEFORE the LAA PE-header rewrite this add-in performs itself.
' A caller that read output.sha256 (the documented "artifact identity") could therefore
' never reproduce it, and nothing in the receipt said the value was a stale snapshot.
'
' The fix lives HERE - one implementation, both build routes call it:
'   * ArtifactFinalSnapshot: one artifact sample taken with a DOUBLE read. Two independent
'     stat+sha256 passes whose agreement is the evidence that the bytes were stable WHILE
'     they were hashed; a locked or short-read file yields NO digest plus a reason, never
'     the digest of an empty input.
'   * FinalizeArtifactOutput: re-shapes a tool's receipt so output{} IS that final sample,
'     moves what output{} used to say out of the way (preRewrite, only when asked for) and
'     turns the ORDERING itself into evidence: hashedAfterRewrites / samplePoint /
'     rewriteSteps[].
'
' WHY hashedAfterRewrites is NOT a constant: it is the AND of three MEASUREMENTS -
' (1) the two reads agreed (stableDuringHash), (2) every rewrite step that was REQUESTED
' completed, (3) no rewrite step reported an error. A failed LAA write, a locked artifact
' or a file that moved while being hashed all answer false - that is what makes the flag
' evidence instead of a promise.
' =====================================================================

Public Function ArtifactFinalSnapshot(ByVal artifactPath As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = artifactPath
    d("exists") = False
    d("size") = 0
    d("mtime") = 0
    d("mtimeText") = ""
    ' a digest that could not be computed is OMITTED with a reason (never the digest of an
    ' empty input - the same F1 rule the older ArtifactFullSnap follows).
    d("sha256") = ""
    d("sha256Available") = False
    d("sha_unavailable_reason") = "artifact_does_not_exist"
    d("sampledAfterRewrites") = True
    d("samplePoint") = "post_rewrite"
    d("samplePointDetail") = "after_compile_and_after_every_post_processing_rewrite"
    d("statReads") = 0
    d("stableDuringHash") = False
    d("rehashNote") = ""
    If Len(artifactPath) = 0 Then
        d("sha_unavailable_reason") = "no_artifact_path"
        Set ArtifactFinalSnapshot = d
        Exit Function
    End If
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(artifactPath) Then
        Err.Clear
        Set ArtifactFinalSnapshot = d
        Exit Function
    End If
    d("exists") = True
    d("statReads") = 2
    Dim sz1 As Double, mt1 As Double, sz2 As Double, mt2 As Double
    Dim why1 As String, why2 As String
    Dim sha1 As String, sha2 As String
    ' READ 1
    sz1 = CDbl(fso.GetFile(artifactPath).Size)
    mt1 = CDbl(fso.GetFile(artifactPath).DateLastModified)
    why1 = ""
    sha1 = FileSha256Why(artifactPath, why1)
    ' READ 2 - the independent second pass
    sz2 = CDbl(fso.GetFile(artifactPath).Size)
    mt2 = CDbl(fso.GetFile(artifactPath).DateLastModified)
    why2 = ""
    sha2 = FileSha256Why(artifactPath, why2)
    d("size") = sz2
    d("mtime") = mt2
    d("mtimeText") = Format$(mt2, "yyyy-mm-dd hh:nn:ss")
    If Len(sha2) > 0 Then
        d("sha256") = sha2
        d("sha256Available") = True
        d("sha_unavailable_reason") = ""
    Else
        d("sha_unavailable_reason") = why2
    End If
    d("stableDuringHash") = ((sz1 = sz2) And (mt1 = mt2) And (Len(sha1) > 0) And (StrComp(sha1, sha2, vbTextCompare) = 0))
    If Not CBool(d("stableDuringHash")) Then
        d("rehashNote") = "the two reads of the artifact did not agree (size/mtime/sha256), so no stable sample can be claimed; read it again before trusting these values"
    End If
    Err.Clear
    Set ArtifactFinalSnapshot = d
End Function

' Re-shape a build receipt so output{} IS the FINAL artifact (see the block comment above).
' publishPreRewrite=false (the default) keeps the pre-rewrite OBJECT unpublished: only its
' digest travels (preRewriteSha256), because a caller that once read a stale object as the
' artifact identity is exactly the defect this function removes.
Public Sub FinalizeArtifactOutput(ByRef root As Object, ByVal artifactPath As String, Optional ByVal publishPreRewrite As Boolean = False)
    If root Is Nothing Then Exit Sub
    Dim laaReq As Boolean, laaSet As Boolean, laaAlready As Boolean
    Dim laaErr As String
    laaReq = DictBool(root, "laaRequested", False)
    laaSet = DictBool(root, "laaSet", False)
    laaAlready = DictBool(root, "laaAlready", False)
    laaErr = DictStr(root, "laaError", "")
    Dim prev As Object
    Set prev = DictObj(root, "output")
    Dim fin As Object
    Set fin = DictObj(root, "final")
    ' FIX-23: an EMPTY artifact path means the caller never resolved an output file (the route
    ' refused before the compile, or the project names no ExeName32). That is not a sample of an
    ' artifact that does not exist - it is the ABSENCE of a subject, and replacing a real
    ' post-compile output{} with a fabricated empty one would erase evidence. So the step is
    ' SKIPPED, structured, and the caller keeps the output{} it published.
    If Len(Trim$(artifactPath)) = 0 Then
        root("outputFinalizeOk") = False
        root("outputFinalizeSkipped") = True
        root("outputFinalizeReason") = "artifact_missing"
        root("artifactMissing") = True
        root("outputFinalizeAdvice") = "no artifact path reached the finalize step (the build never resolved an output file), so output{} was left exactly as the caller published it and is NOT a final post-rewrite sample; resolve ExeName32/output_path, run the compile, then re-read artifact{}"
        root("hashedAfterRewrites") = False
        root("samplePoint") = "not_sampled"
        root("samplePointDetail") = "no_artifact_path_to_sample"
        Exit Sub
    End If
    Dim out As Object
    Set out = ArtifactFinalSnapshot(artifactPath)
    ' FIX-23: a path that was resolved but produced NO file is a real MEASURED state - publish
    ' that measurement (never a stale compile-time sample) and say what to do next.
    If Not CBool(out("exists")) Then
        root("artifactMissing") = True
        root("outputFinalizeAdvice") = "the finalize step sampled " & artifactPath & " and the file does not exist; output{} is that measured absence (exists=false), not a stale pre-compile sample. Read verdict/buildState/firstError and the output path, then rebuild."
    End If
    ' ---- rewriteSteps[]: what ran BETWEEN the compile and this sample --------------
    Dim steps As Object
    Dim st As Object
    Set steps = New Collection
    Set st = CreateObject("Scripting.Dictionary")
    st("step") = "laa_pe_header_patch"
    st("requested") = laaReq
    st("applied") = laaSet
    st("alreadySet") = laaAlready
    st("completed") = ((laaReq = False) Or laaSet Or laaAlready)
    st("error") = laaErr
    Dim attempted As Long, completed As Long, failed As Long
    attempted = 0
    completed = 0
    failed = 0
    If laaReq Then
        ' the step WAS requested: completed means it either patched the header or found the
        ' flag already set; a non-empty error means it did not do either.
        attempted = 1
        If laaSet Or laaAlready Then completed = 1
        If Len(laaErr) > 0 Then failed = 1
    End If
    ' ---- did the post-processing actually MOVE the artifact bytes? ------------------
    Dim prevSha As String
    Dim moved As Boolean
    prevSha = DictStr(prev, "sha256", "")
    moved = False
    If (Len(prevSha) > 0) And CBool(out("sha256Available")) Then
        moved = (StrComp(prevSha, CStr(out("sha256")), vbTextCompare) <> 0)
    End If
    If moved Then
        st("bytesEffect") = "artifact_bytes_changed_after_compile"
    ElseIf Len(prevSha) = 0 Then
        st("bytesEffect") = "no_pre_rewrite_digest_available"
    Else
        st("bytesEffect") = "artifact_bytes_unchanged_since_compile"
    End If
    steps.Add st
    ' ---- hashedAfterRewrites: MEASURED, never a constant ---------------------------
    Dim hAfter As Boolean
    hAfter = CBool(out("stableDuringHash")) And (completed = attempted) And (failed = 0)
    If Not CBool(out("exists")) Then hAfter = False
    ' the PE facts a caller uses to prove the LAA write landed are taken from final{} when
    ' the caller published one, so output{} and final{} cannot contradict each other.
    If Not fin Is Nothing Then
        If DictHas(fin, "peFlags") Then out("peFlags") = fin("peFlags")
        If DictHas(fin, "laaFlag") Then out("laaFlag") = fin("laaFlag")
        If DictHas(fin, "verifiedOnDisk") Then out("verifiedOnDisk") = fin("verifiedOnDisk")
    End If
    ' ---- the receipt ---------------------------------------------------------------
    If Not prev Is Nothing Then
        root("preRewriteSha256") = prevSha
        root("preRewriteNote") = "preRewriteSha256 is the artifact as it looked BEFORE the post-processing rewrites (right after the compile). It is a COMPARISON POINT only - never the artifact identity; output{} is the identity."
        If publishPreRewrite Then
            Set root("preRewrite") = prev
        End If
    End If
    Set root("output") = out
    root("hashedAfterRewrites") = hAfter
    root("samplePoint") = CStr(out("samplePoint"))
    root("samplePointDetail") = CStr(out("samplePointDetail"))
    root("outputSampleNote") = "output{size,mtime,sha256} is sampled AFTER the compile and AFTER every step listed in rewriteSteps[]; compare it with an external Get-FileHash of the same file"
    Set root("rewriteSteps") = steps
    root("rewriteStepsAttempted") = attempted
    root("rewriteStepsCompleted") = completed
    root("rewriteStepsFailed") = failed
    root("artifactBytesChangedByRewrite") = moved
    root("artifactHashStable") = CBool(out("stableDuringHash"))
End Sub

' ---- small Dictionary readers that never auto-add a key (reading a missing key through
' ---- the default-property syntax ADDS it as Empty - the known Scripting.Dictionary trap)
Private Function DictHas(ByVal d As Object, ByVal k As String) As Boolean
    If d Is Nothing Then Exit Function
    DictHas = d.Exists(k)
End Function

Private Function DictStr(ByVal d As Object, ByVal k As String, ByVal def As String) As String
    DictStr = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(k) Then Exit Function
    If IsObject(d.Item(k)) Then Exit Function
    Dim v As Variant
    v = d.Item(k)
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    DictStr = CStr(v)
End Function

Private Function DictBool(ByVal d As Object, ByVal k As String, ByVal def As Boolean) As Boolean
    DictBool = def
    If d Is Nothing Then Exit Function
    If Not d.Exists(k) Then Exit Function
    If IsObject(d.Item(k)) Then Exit Function
    Dim v As Variant
    v = d.Item(k)
    If IsNull(v) Then Exit Function
    If IsObject(v) Then Exit Function
    If VarType(v) = vbBoolean Then
        DictBool = CBool(v)
    Else
        DictBool = (LCase$(CStr(v)) = "true") Or (CStr(v) = "1")
    End If
End Function

Private Function DictObj(ByVal d As Object, ByVal k As String) As Object
    ' FIX-23 - MEASURED, not inferred (isolated VB6 /make probe, fix/_dict450):
    ' reading an OBJECT item Set-less (v = d.Item(k)) raises runtime error 450 "wrong number of
    ' arguments or invalid property assignment", because a Variant assignment resolves the item's
    ' DEFAULT member and Scripting.Dictionary's default member is Item(Key), which cannot be
    ' called with no argument. DictStr/DictBool above already guard with IsObject BEFORE the read;
    ' this one did not. FinalizeArtifactOutput passes output{} and final{} - BOTH Dictionaries -
    ' so it raised 450 and a COMPLETE build receipt became reason=unexpected_error(vbErrorCode:450).
    ' Probe evidence: asIs_err=450 asIs_got=nothing | fixed_err=0 fixed_got=object.
    If d Is Nothing Then Exit Function
    If Not d.Exists(k) Then Exit Function
    If IsObject(d.Item(k)) Then Set DictObj = d.Item(k)
End Function
