Attribute VB_Name = "Build"
Option Explicit

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
    encUsed = Encoding.DetectEncoding(b)
    If encUsed = "empty" Then encUsed = "ansi"
    ' utf-16be has no decoder here; decoding it as ANSI at least keeps the byte order
    ' explicit instead of silently producing mojibake.
    If encUsed = "utf-16be" Then encUsed = "ansi"
    Encoding.DecodeBytes b, encUsed, out
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
            For b2 = a + 1 To UBound(lines)
                MaybePump
                If lines(b2) < lines(a) Then
                    tmp = lines(a): lines(a) = lines(b2): lines(b2) = tmp
                End If
            Next
        Next
        For i = LBound(lines) To UBound(lines)
            MaybePump
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
                End Select
            End If
        End If
    Loop
    ts.Close
    Set c("comps") = comps
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
