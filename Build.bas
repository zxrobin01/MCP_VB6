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
    ' Card TASK-P1-007 (i18n): product-first verdict fields
    artifactSize As Double
    okByKeyword As Boolean
    verdict As String
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' ---- Card TASK-P1-015: path identity ---------------------------------------
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
    ' TASK-P1-126 (F2): the old body ALWAYS decoded as ANSI. A log that is really
    ' UTF-8 (or UTF-16) then came back as U+FFFD replacement characters, so the
    ' compiler message ("syntax error" / the localized equivalent) was unreadable.
    ' ReadLogAny reads it in ITS OWN encoding and reports which one was used.
    Dim s As String, e As String
    ReadLogAny logPath, s, e
    ReadLogAnsi = s
End Function

' TASK-P1-126 (F2): detect the log's real encoding, decode with it, and return the
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

Public Function ErrorKey(ByVal msg As String) As String
    ' Card TASK-P1-007 (i18n): the localized words are kept, but English keywords
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
    Dim r As TBuildResult, lines() As String, i As Long, ln As String, t As String
    r.ok = False
    r.exitCode = -999
    r.headerOffset = headerOffset
    r.raw = ReadLogAnsi(logPath)
    t = r.raw
    r.okByKeyword = (InStr(t, COk()) > 0)
    If (Not r.okByKeyword) And Len(Trim$(t)) > 0 Then
        If InStr(1, t, "success", vbTextCompare) > 0 Then r.okByKeyword = True
    End If
    r.ok = r.okByKeyword
    lines = Split(Replace(Replace(t, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        MaybePump
        ln = lines(i)
        If InStr(ln, CErr()) > 0 Or InStr(1, ln, "compile error", vbTextCompare) > 0 Then
            r.file = BetweenQuotes(ln)
            r.reportedLine = LineAfter(ln)
            r.message = MsgAfter(ln)
            r.key = ErrorKey(r.message)
            If headerOffset >= 0 And r.reportedLine > 0 Then r.diskLine = r.reportedLine + headerOffset + 1
            Exit For
        End If
    Next
    ParseBuild = r
End Function

' TASK-P1-125 (MCP-18): how many compile errors does a VB6 /out log carry?
' One build stops at the FIRST error, but a caller still needs a count to judge
' whether the log is a single failure or a cascade; the count is measured on the
' same lines ParseBuild recognises, never assumed to be 1.
Public Function CountBuildErrors(ByVal raw As String) As Long
    If Len(raw) = 0 Then Exit Function
    Dim lines() As String, i As Long, n As Long
    lines = Split(Replace(Replace(raw, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(lines)
        If InStr(lines(i), CErr()) > 0 Or InStr(1, lines(i), "compile error", vbTextCompare) > 0 Then n = n + 1
    Next
    CountBuildErrors = n
End Function

' TASK-P1-125 (MCP-18): the compiler's line numbers are CODE-PANE lines. When the
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
' Card TASK-P1-007 (i18n): product-first build verdict.
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
    ' Card TASK-P1-015 (found while building the junction fixture): TextStream.ReadLine
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
' Card TASK-P1-008: incremental build feedback.
' Fingerprint = sha256(.vbp bytes) + sorted sha256 of every component file the
' .vbp references, plus /d defines. Disk based and language independent.
' ---------------------------------------------------------------------------
Public Function FileSha256(ByVal path As String) As String
    ' TASK-P1-126 (F1): delegate; the WHY-carrying form below is what lets a caller
    ' distinguish "no digest" from "the digest of an empty input".
    Dim why As String
    why = ""
    FileSha256 = FileSha256Why(path, why)
End Function

' TASK-P1-126 (F1): a sha256 that can NEVER silently be the digest of an EMPTY input.
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
    ' Card TASK-P1-015: the cache key must be the project IDENTITY, not its spelling.
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
