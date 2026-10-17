Attribute VB_Name = "ProjShip"
Option Explicit

' ---------------------------------------------------------------------------
' Project group (.vbg), project-level snapshots/rollback, delivery packaging
' and captured subprocess execution (for read-only VCS queries).
'
' Everything here is file/text level: no IDE object model is required, so these
' tools also work on a project that is not open in the IDE.
' ---------------------------------------------------------------------------

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' =========================================================== .vbg project group

' Lists the projects referenced by a .vbg file.
Public Function VbgListJson(ByVal vbgPath As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = vbgPath
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    d("exists") = fso.FileExists(vbgPath)
    Dim items As Object
    Set items = New Collection
    If fso.FileExists(vbgPath) Then
        Dim lines() As String, i As Long, p As String
        lines = SplitLines(ReadTextFile(vbgPath))
        For i = LBound(lines) To UBound(lines)
            MaybePump
            p = VbgProjectOf(lines(i))
            If Len(p) > 0 Then
                Dim absP As String
                absP = AbsPath(vbgPath, p)
                Dim e As Object
                Set e = CreateObject("Scripting.Dictionary")
                e("project") = p
                e("resolved") = absP
                e("exists") = fso.FileExists(absP)
                items.Add e
            End If
        Next
    End If
    Set d("projects") = items
    d("count") = items.Count
    VbgListJson = JsonStringify(d)
End Function

' Adds a project to a .vbg (idempotent). Creates the file when missing.
Public Function VbgAddJson(ByVal vbgPath As String, ByVal projectPath As String) As String
    If Len(vbgPath) = 0 Then Err.Raise vbObjectError + 71, "ProjShip", "path is required (the .vbg file)"
    If Len(projectPath) = 0 Then Err.Raise vbObjectError + 71, "ProjShip", "project is required (the .vbp to add)"
    Dim text As String
    text = ReadTextFile(vbgPath)
    Dim eol As String
    eol = LineEndingOf(text)
    If Len(text) = 0 Then text = "Microsoft Visual Basic 6.0 Project Group" & eol
    Dim lines() As String, i As Long, p As String
    lines = SplitLines(text)
    Dim already As Boolean
    already = False
    For i = LBound(lines) To UBound(lines)
        MaybePump
        p = VbgProjectOf(lines(i))
        If Len(p) > 0 Then
            If StrComp(NormPath(p), NormPath(projectPath), vbTextCompare) = 0 Then already = True
        End If
    Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    If already Then
        d("added") = False
        d("reason") = "already present"
    Else
        BackupFile vbgPath
        Dim newLine As String
        If InStr(projectPath, " ") > 0 Then
            newLine = "Project=" & Chr$(34) & projectPath & Chr$(34)
        Else
            newLine = "Project=" & projectPath
        End If
        If Not WriteTextFile(vbgPath, TrimRight$(Join(lines, eol)) & eol & newLine & eol) Then
        VbgAddJson = ShipWriteRefusal("vbg_group", vbgPath, "write_not_visible", "the .vbg could not be read back after writing; nothing is claimed to have changed")
        Exit Function
    End If
        d("added") = True
    End If
    d("path") = vbgPath
    d("project") = projectPath
    VbgAddJson = JsonStringify(d)
End Function

' Removes a project from a .vbg.
Public Function VbgRemoveJson(ByVal vbgPath As String, ByVal projectPath As String) As String
    Dim text As String
    text = ReadTextFile(vbgPath)
    If Len(text) = 0 Then Err.Raise vbObjectError + 72, "ProjShip", "project group not found or empty: " & vbgPath
    Dim eol As String
    eol = LineEndingOf(text)
    Dim lines() As String, i As Long, p As String, removed As Long, out As String
    lines = SplitLines(text)
    removed = 0
    out = ""
    For i = LBound(lines) To UBound(lines)
        MaybePump
        p = VbgProjectOf(lines(i))
        If Len(p) > 0 And StrComp(NormPath(p), NormPath(projectPath), vbTextCompare) = 0 Then
            removed = removed + 1
        Else
            If Len(out) > 0 Then out = out & eol
            out = out & lines(i)
        End If
    Next
    If removed > 0 Then
        BackupFile vbgPath
        If Not WriteTextFile(vbgPath, out & eol) Then
        VbgRemoveJson = ShipWriteRefusal("vbg_group", vbgPath, "write_not_visible", "the .vbg could not be read back after writing; nothing is claimed to have changed")
        Exit Function
    End If
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = vbgPath
    d("removed") = removed
    VbgRemoveJson = JsonStringify(d)
End Function

Private Function VbgProjectOf(ByVal ln As String) As String
    Dim t As String
    t = Trim$(ln)
    VbgProjectOf = ""
    If Len(t) < 9 Then Exit Function
    If StrComp(Left$(t, 8), "Project=", vbTextCompare) <> 0 Then Exit Function
    t = Trim$(Mid$(t, 9))
    If Len(t) >= 2 Then
        If Left$(t, 1) = Chr$(34) And Right$(t, 1) = Chr$(34) Then t = Mid$(t, 2, Len(t) - 2)
    End If
    VbgProjectOf = t
End Function

' =========================================================== project snapshot

' Copies the .vbp plus every component file into a snapshot directory and
' writes snapshot.json (original path, sha256 and size per file).
Public Function SnapshotCreateJson(ByVal vbpPath As String, ByVal dir As String) As String
    If Len(vbpPath) = 0 Then Err.Raise vbObjectError + 73, "ProjShip", "path is required (the .vbp to snapshot)"
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(vbpPath) Then Err.Raise vbObjectError + 73, "ProjShip", "project file not found: " & vbpPath
    If Len(dir) = 0 Then dir = fso.BuildPath(DirOf(vbpPath), "_vb6mcp_snapshots\" & Stamp())
    EnsureDir dir
    Dim files As Object
    Set files = New Collection
    Dim saved As Object
    Set saved = CreateObject("Scripting.Dictionary")
    CopyInto files, saved, fso, vbpPath, dir
    Dim comps As Object
    Set comps = ComponentPaths(vbpPath)
    Dim v As Variant
    For Each v In comps
        If fso.FileExists(CStr(v)) Then
            If Not saved.Exists(LCase$(fso.GetFileName(CStr(v)))) Then
                CopyInto files, saved, fso, CStr(v), dir
            End If
        End If
    Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("dir") = dir
    d("vbp") = vbpPath
    d("createdAt") = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    d("fileCount") = files.Count
    d("verified") = "every entry below exists on disk with the recorded sha256 (and, when srcSha256 is present, equals its source file)"
    Set d("files") = files
    Set d("references") = ReferencePaths(vbpPath)
    Dim s As String
    s = JsonStringify(d)
    Dim manPath As String
    manPath = fso.BuildPath(dir, "snapshot.json")
    If Not WriteTextFile(manPath, s) Then
        SnapshotCreateJson = ShipWriteRefusal("project_snapshot", manPath, "write_not_visible", "the snapshot manifest could not be read back after writing; nothing is claimed to be snapshotted")
        Exit Function
    End If
    ' verify EVERY file the manifest records - exists + the recorded sha256.
    Dim vf As Variant, vName As String, vPath As String, bad As String
    Dim unver As String, nUnver As Long
    bad = ""
    unver = ""
    nUnver = 0
    For Each vf In files
        vName = CStr(vf("name"))
        vPath = fso.BuildPath(dir, vName)
        If Not fso.FileExists(vPath) Then
            bad = vName & " (recorded in the manifest but NOT on disk)"
            Exit For
        End If
        If vf.Exists("sha256") Then
            If Sha256HexFile(vPath) <> CStr(vf("sha256")) Then
                bad = vName & " (on disk but its sha256 differs from the manifest)"
                Exit For
            End If
            ' the copy must also equal the SOURCE it was made from
            If vf.Exists("srcSha256") Then
                If Sha256HexFile(vPath) <> CStr(vf("srcSha256")) Then
                    bad = vName & " (the snapshot copy differs from its source file)"
                    Exit For
                End If
            End If
        Else
            ' remaining: an entry with NO recorded sha256 is NAMED here - a silent
            ' skip would let an unverifiable entry pass as a verified snapshot.
            nUnver = nUnver + 1
            If Len(unver) > 0 Then unver = unver & ","
            unver = unver & vName
        End If
    Next
    If Len(bad) > 0 Then
        SnapshotCreateJson = ShipWriteRefusal("project_snapshot", dir, "snapshot_incomplete", "the snapshot is NOT usable: " & bad)
        Exit Function
    End If
    If nUnver > 0 Then
        SnapshotCreateJson = ShipWriteRefusal("project_snapshot", dir, "snapshot_unverified", CStr(nUnver) & " manifest entr(y/ies) carry no sha256 and were therefore NOT verified: " & unver)
        Exit Function
    End If
    SnapshotCreateJson = s
End Function

' Restores every file recorded in snapshot.json (each overwritten file gets .bak).
Public Function SnapshotRestoreJson(ByVal dir As String, ByRef restoredCsv As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim man As String
    man = fso.BuildPath(dir, "snapshot.json")
    If Not fso.FileExists(man) Then Err.Raise vbObjectError + 74, "ProjShip", "snapshot.json not found under: " & dir
    Dim text As String
    text = ReadTextFile(man)
    Dim restored As Long, skipped As Long
    restored = 0
    skipped = 0
    restoredCsv = ""
    ' remaining: copies that did NOT survive a read-back.
    Dim failed As Long
    Dim failedCsv As String
    failed = 0
    failedCsv = ""
    ' remaining: snapshot files whose bytes no longer match the manifest.
    Dim corrupt As Long
    Dim corruptCsv As String
    corrupt = 0
    corruptCsv = ""
    Dim cur As Long, endPos As Long, chunk As String, src As String, nm As String
    cur = 1
    Do
        cur = InStr(cur, text, "{")
        If cur = 0 Then Exit Do
        endPos = InStr(cur, text, "}")
        If endPos = 0 Then Exit Do
        chunk = Mid$(text, cur, endPos - cur + 1)
        src = JsonFieldOf(chunk, "src")
        nm = JsonFieldOf(chunk, "name")
        If Len(src) > 0 And Len(nm) > 0 Then
            Dim from As String
            from = fso.BuildPath(dir, nm)
            If fso.FileExists(from) Then
                ' remaining: restore ONLY from a file that still matches the manifest, so
                ' a snapshot whose bytes were changed after it was taken cannot silently become
                ' the project's source. The mismatch is NAMED in the receipt.
                Dim mSha As String
                mSha = JsonFieldOf(chunk, "sha256")
                If Len(mSha) = 64 Then
                    If Sha256HexFile(from) <> mSha Then
                        corrupt = corrupt + 1
                        If Len(corruptCsv) > 0 Then corruptCsv = corruptCsv & ","
                        corruptCsv = corruptCsv & nm
                        cur = endPos + 1
                        GoTo NextChunk
                    End If
                End If
                If fso.FileExists(src) Then fso.CopyFile src, src & ".bak", True
                EnsureDir DirOf(src)
                fso.CopyFile from, src, True
                ' remaining: a file counts as RESTORED only when the bytes are
                ' readable back at the destination; otherwise it is reported as FAILED.
                Dim rstB() As Byte
                ReadAllBytesTo src, rstB
                If ByteCount(rstB) > 0 Then
                    restored = restored + 1
                    If Len(restoredCsv) > 0 Then restoredCsv = restoredCsv & ","
                    restoredCsv = restoredCsv & nm
                Else
                    failed = failed + 1
                    If Len(failedCsv) > 0 Then failedCsv = failedCsv & ","
                    failedCsv = failedCsv & nm
                End If
            Else
                skipped = skipped + 1
            End If
        End If
        cur = endPos + 1
NextChunk:
    Loop
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("dir") = dir
    d("restored") = restored
    d("skipped") = skipped
    d("files") = restoredCsv
    d("failed") = failed
    d("failedFiles") = failedCsv
    d("corrupt") = corrupt
    d("corruptFiles") = corruptCsv
    If corrupt > 0 Then
        d("ok") = False
        d("reason") = "snapshot_corrupt"
        d("note") = "one or more snapshot files no longer match snapshot.json (they were changed or truncated after the snapshot was taken); those files were NOT restored and are named in corruptFiles."
    End If
    If failed > 0 Then
        d("ok") = False
        d("reason") = "restore_not_readable_back"
        d("note") = "one or more files were copied back but are NOT readable at the destination: they are counted in failed/failedFiles, never in restored."
    End If
    SnapshotRestoreJson = JsonStringify(d)
End Function

' Lists the snapshot directories that sit next to a project.
Public Function SnapshotListJson(ByVal vbpPath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim root As String
    root = fso.BuildPath(DirOf(vbpPath), "_vb6mcp_snapshots")
    Dim items As Object
    Set items = New Collection
    If fso.FolderExists(root) Then
        Dim f As Object
        For Each f In fso.GetFolder(root).SubFolders
            Dim e As Object
            Set e = CreateObject("Scripting.Dictionary")
            e("name") = f.Name
            e("dir") = f.Path
            Dim man As String
            man = fso.BuildPath(f.Path, "snapshot.json")
            If fso.FileExists(man) Then e("createdAt") = JsonFieldOf(ReadTextFile(man), "createdAt")
            items.Add e
        Next
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("root") = root
    d("count") = items.Count
    Set d("snapshots") = items
    SnapshotListJson = JsonStringify(d)
End Function

' =========================================================== delivery package

' Assembles a delivery folder: the built exe, an external manifest, the
' dependency list taken from the .vbp Reference= lines, and install.bat.
Public Function PackageJson(ByVal vbpPath As String, ByVal exePath As String, ByVal outDir As String, ByVal level As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(vbpPath) = 0 Then Err.Raise vbObjectError + 75, "ProjShip", "path is required (the .vbp to package)"
    If Not fso.FileExists(vbpPath) Then Err.Raise vbObjectError + 75, "ProjShip", "project file not found: " & vbpPath
    Dim text As String
    text = ReadTextFile(vbpPath)
    If Len(exePath) = 0 Then exePath = fso.BuildPath(DirOf(vbpPath), VbpValue(text, "ExeName32", "app.exe"))
    If Len(outDir) = 0 Then outDir = fso.BuildPath(DirOf(vbpPath), "_vb6mcp_dist")
    EnsureDir outDir
    Dim made As Object
    Set made = New Collection
    Dim exeOk As Boolean
    exeOk = fso.FileExists(exePath)
    If exeOk Then
        fso.CopyFile exePath, fso.BuildPath(outDir, fso.GetFileName(exePath)), True
        made.Add fso.GetFileName(exePath)
        Dim man As String
        man = ProjMeta.ManifestText(level, True, "permonitorv2", False, fso.GetFileName(exePath))
        Dim wFail As String
        If Not WriteTextFile(fso.BuildPath(outDir, fso.GetFileName(exePath) & ".manifest"), man) Then wFail = fso.GetFileName(exePath) & ".manifest"
        made.Add fso.GetFileName(exePath) & ".manifest"
    End If
    Dim deps As Object
    Set deps = ReferencePaths(vbpPath)
    Dim listed As Object
    Set listed = New Collection
    Dim regBat As String
    ' the header is composed AFTER the loop (it needs the dependency count), so
    ' the per-dependency lines go into their own buffer first.
    Dim regBody As String, regN As Long
    regBody = ""
    regN = 0
    regBat = ""
    Dim missing As Object
    Set missing = New Collection
    Dim v As Variant, dp As String, ext As String
    For Each v In deps
        dp = CStr(v)
        ext = LCase$(fso.GetExtensionName(dp))
        Dim exists As Boolean
        exists = fso.FileExists(dp)
        Dim e As Object
        Set e = CreateObject("Scripting.Dictionary")
        e("path") = dp
        e("name") = fso.GetFileName(dp)
        e("exists") = exists
        If exists Then e("sha256") = Sha256HexFile(dp)
        listed.Add e
        If Not exists Then missing.Add dp
        If exists And (ext = "ocx" Or ext = "dll") Then
            ' NEVER a bare regsvr32. A bare name resolves through PATH and can hit a
            ' stale SECOND installation of Windows (measured on a host whose SystemRoot is NOT on
            ' the system drive, while a leftover Windows tree sits there) where EVERY 32-bit registration fails with 0xC0000139 -
            ' and /s prints nothing, so the failure is silent. So: the SysWOW64 copy of THIS
            ' installation, never a hardcoded C:\Windows. The exit code IS acted on here: the rc
            ' of a direct call is trustworthy (the stale-0 trap was a PowerShell $LASTEXITCODE
            ' artifact of the CALLER, not something regsvr32 does).
            regBody = regBody & "rem dependency: " & fso.GetFileName(dp) & vbCrLf
            regBody = regBody & Chr$(34) & "%REGSVR%" & Chr$(34) & " /s " & Chr$(34) & fso.GetFileName(dp) & Chr$(34) & vbCrLf
            regBody = regBody & "if errorlevel 1 goto regfail" & vbCrLf
            regN = regN + 1
        End If
    Next
    If Not exeOk Then missing.Add exePath
    ' hardened install.bat. Header + existence gate first, then one rc-checked
    ' registration per dependency, then an explicit non-zero exit on any refusal.
    regBat = "@echo off" & vbCrLf
    regBat = regBat & "rem Generated by VB6MCP package_app" & vbCrLf
    regBat = regBat & "rem a 32-bit OCX/DLL must be registered with the SysWOW64 regsvr32 of" & vbCrLf
    regBat = regBat & "rem THIS Windows installation. A bare regsvr32 resolves through PATH and can hit a" & vbCrLf
    regBat = regBat & "rem stale second installation (SystemRoot not on the system drive, a leftover tree there) where" & vbCrLf
    regBat = regBat & "rem every 32-bit registration fails with 0xC0000139 - silently, because /s prints nothing." & vbCrLf
    regBat = regBat & "setlocal EnableExtensions" & vbCrLf
    regBat = regBat & "set ""REGSVR=%SystemRoot%\SysWOW64\regsvr32.exe""" & vbCrLf
    regBat = regBat & "if not exist ""%REGSVR%"" goto nocom" & vbCrLf
    regBat = regBat & "set ""DEPS=" & CStr(regN) & """" & vbCrLf
    regBat = regBat & regBody
    regBat = regBat & "echo install.bat: %DEPS% dependency file(s) registered." & vbCrLf
    regBat = regBat & "endlocal & exit /b 0" & vbCrLf
    regBat = regBat & ":nocom" & vbCrLf
    regBat = regBat & ">&2 echo INSTALL FAILED: 32-bit regsvr32 is missing: ""%REGSVR%"" (SystemRoot=%SystemRoot%)." & vbCrLf
    regBat = regBat & ">&2 echo   This package registers 32-bit OCX/DLL dependencies; the SysWOW64 regsvr32 of" & vbCrLf
    regBat = regBat & ">&2 echo   THIS Windows installation is required, and a bare regsvr32 is deliberately NOT used." & vbCrLf
    regBat = regBat & "endlocal & exit /b 2" & vbCrLf
    regBat = regBat & ":regfail" & vbCrLf
    regBat = regBat & ">&2 echo INSTALL FAILED: a dependency above was NOT registered (regsvr32 returned non-zero)." & vbCrLf
    regBat = regBat & ">&2 echo   Identify it by the 'rem dependency:' line above the last regsvr32 line, then re-run" & vbCrLf
    regBat = regBat & ">&2 echo   that command by hand and read its error; a 32-bit DLL must use ""%REGSVR%""." & vbCrLf
    regBat = regBat & "endlocal & exit /b 1" & vbCrLf
    If Not WriteTextFile(fso.BuildPath(outDir, "install.bat"), regBat) Then wFail = "install.bat"
    made.Add "install.bat"
    Dim depTxt As String
    depTxt = "Dependencies (from the .vbp Reference= lines)" & vbCrLf
    For Each v In listed
        depTxt = depTxt & v("name") & vbTab & v("path") & vbTab & VbFlag(CBool(v("exists")), "present", "MISSING") & vbCrLf
    Next
    If Not WriteTextFile(fso.BuildPath(outDir, "dependencies.txt"), depTxt) Then wFail = "dependencies.txt"
    made.Add "dependencies.txt"
    Dim readme As String
    readme = "VB6MCP delivery package" & vbCrLf
    readme = readme & "Generated: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbCrLf
    readme = readme & "Project  : " & vbpPath & vbCrLf
    readme = readme & "Exe      : " & exePath & vbCrLf & vbCrLf
    readme = readme & "The external manifest (" & fso.GetFileName(exePath) & ".manifest) must stay next to the exe." & vbCrLf
    readme = readme & "Run install.bat on the target machine to register the OCX/DLL dependencies." & vbCrLf
    readme = readme & "install.bat uses %SystemRoot%\SysWOW64\regsvr32.exe (never a bare regsvr32, never a hardcoded C:\Windows)" & vbCrLf
    readme = readme & "and exits NON-ZERO when a dependency cannot be registered or when that regsvr32 is missing." & vbCrLf
    If Not WriteTextFile(fso.BuildPath(outDir, "README.txt"), readme) Then wFail = "README.txt"
    made.Add "README.txt"
    ' a write that could not be read back is refused AT ITS OWN SITE, not merely covered
    ' by the artifact sweep below (that would put the burden back on 'the caller remembers').
    If Len(wFail) > 0 Then
        PackageJson = ShipWriteRefusal("package_app", fso.BuildPath(outDir, wFail), "write_not_visible", "an artifact could not be read back after writing; nothing is claimed to have been packaged")
        Exit Function
    End If
    ' every artifact this function CLAIMS to have made must be on disk and non-empty.
    Dim va As Variant, aName As String, aPath As String, aBad As String
    aBad = ""
    For Each va In made
        aName = CStr(va)
        aPath = fso.BuildPath(outDir, aName)
        If Not fso.FileExists(aPath) Then
            aBad = aName & " (claimed but NOT on disk)"
            Exit For
        End If
        If fso.GetFile(aPath).Size <= 0 Then
            aBad = aName & " (exists but is EMPTY)"
            Exit For
        End If
    Next
    If Len(aBad) > 0 Then
        PackageJson = ShipWriteRefusal("package_app", outDir, "package_incomplete", "the delivery folder is NOT complete: " & aBad)
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("outDir") = outDir
    d("exePath") = exePath
    d("exeFound") = exeOk
    d("level") = level
    Set d("artifacts") = made
    Set d("dependencies") = listed
    Set d("missing") = missing
    PackageJson = JsonStringify(d)
End Function

' =========================================================== captured process

' Runs a process and returns its captured stdout: {"exe","args","exitCode",
' "stdout","timedOut"}. Output is buffered, so this suits short queries
' (VCS status/diff), not long streaming jobs.
Public Function RunCaptureJson(ByVal exePath As String, ByVal args As String, ByVal timeoutMs As Long) As String
    If Len(exePath) = 0 Then Err.Raise vbObjectError + 76, "ProjShip", "exe is required"
    If timeoutMs <= 0 Then timeoutMs = 30000
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("exe") = exePath
    d("args") = args
    d("stdout") = ""
    d("timedOut") = False
    d("exitCode") = -1
    On Error GoTo fail
    Dim sh As Object, ex As Object, waited As Long
    Set sh = CreateObject("WScript.Shell")
    Set ex = sh.Exec(Chr$(34) & exePath & Chr$(34) & " " & args)
    waited = 0
    Do While ex.Status = 0
        Sleep 50
        waited = waited + 50
        If waited >= timeoutMs Then
            ex.Terminate
            d("timedOut") = True
            RunCaptureJson = JsonStringify(d)
            Exit Function
        End If
    Loop
    d("exitCode") = ex.ExitCode
    On Error Resume Next
    d("stdout") = ex.StdOut.ReadAll
    On Error GoTo 0
    RunCaptureJson = JsonStringify(d)
    Exit Function
fail:
    d("error") = Err.Description
    RunCaptureJson = JsonStringify(d)
End Function

' =========================================================== helpers

Private Function VbFlag(ByVal b As Boolean, ByVal t As String, ByVal f As String) As String
    If b Then VbFlag = t Else VbFlag = f
End Function

Private Sub CopyInto(ByVal files As Object, ByVal seen As Object, ByVal fso As Object, ByVal src As String, ByVal dir As String)
    Dim nm As String
    nm = fso.GetFileName(src)
    Dim dstP As String
    dstP = fso.BuildPath(dir, nm)
    ' remaining: a FAILED copy must become a named entry with an error, never a
    ' raised run-time error - otherwise the structured snapshot_incomplete refusal is lost.
    Dim cpErr As Long
    cpErr = 0
    On Error Resume Next
    Err.Clear
    fso.CopyFile src, dstP, True
    cpErr = Err.Number
    Err.Clear
    On Error GoTo 0
    ' remaining: the OLD code recorded the SOURCE's size/sha while claiming they
    ' belonged to the packaged copy, and marked the file saved before proving it landed.
    ' Read the DESTINATION back instead.
    Dim ciB() As Byte
    ReadAllBytesTo dstP, ciB
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    e("name") = nm
    e("src") = src
    If cpErr <> 0 Then
        e("error") = "copy_failed"
        e("errorDetail") = CStr(cpErr)
        e("errorTarget") = dstP
    ElseIf ByteCount(ciB) <= 0 Then
        e("error") = "copy_not_readable_back"
        e("errorTarget") = dstP
    Else
        seen(LCase$(nm)) = True
        e("bytes") = fso.GetFile(dstP).Size
        e("sha256") = Sha256HexFile(dstP)
        ' the SOURCE sha is recorded too, so create() can also prove the copy is faithful
        e("srcSha256") = Sha256HexFile(src)
    End If
    files.Add e
End Sub

' Absolute paths of the component files declared in a .vbp.
Public Function ComponentPaths(ByVal vbpPath As String) As Object
    Dim items As Object
    Set items = New Collection
    Dim lines() As String, i As Long, t As String, p As String, semi As Long
    lines = SplitLines(ReadTextFile(vbpPath))
    For i = LBound(lines) To UBound(lines)
        MaybePump
        t = Trim$(lines(i))
        p = ""
        If StartsWith(t, "Module=") Then
            p = Mid$(t, 8)
        ElseIf StartsWith(t, "Class=") Then
            p = Mid$(t, 7)
        ElseIf StartsWith(t, "Form=") Then
            p = Mid$(t, 6)
        ElseIf StartsWith(t, "UserControl=") Then
            p = Mid$(t, 13)
        ElseIf StartsWith(t, "RelatedDoc=") Then
            p = Mid$(t, 12)
        ElseIf StartsWith(t, "ResFile=") Then
            p = Mid$(t, 9)
        End If
        If Len(p) > 0 Then
            semi = InStr(p, ";")
            If semi > 0 Then p = Trim$(Mid$(p, semi + 1))
            If Len(p) > 0 Then items.Add AbsPath(vbpPath, p)
        End If
    Next
    Set ComponentPaths = items
End Function

' File paths found in the Reference= lines of a .vbp.
Public Function ReferencePaths(ByVal vbpPath As String) As Object
    Dim items As Object
    Set items = New Collection
    Dim lines() As String, i As Long, t As String, parts() As String, p As String
    lines = SplitLines(ReadTextFile(vbpPath))
    For i = LBound(lines) To UBound(lines)
        MaybePump
        t = Trim$(lines(i))
        If StartsWith(t, "Reference=") Then
            parts = Split(Mid$(t, 11), "#")
            If UBound(parts) >= 3 Then
                p = Trim$(parts(3))
                If Len(p) > 0 Then items.Add p
            End If
        End If
    Next
    Set ReferencePaths = items
End Function

Private Function VbpValue(ByVal text As String, ByVal key As String, ByVal def As String) As String
    Dim lines() As String, i As Long, t As String, v As String
    lines = SplitLines(text)
    For i = LBound(lines) To UBound(lines)
        MaybePump
        t = Trim$(lines(i))
        If StartsWith(t, key & "=") Then
            v = Trim$(Mid$(t, Len(key) + 2))
            If Len(v) >= 2 Then
                If Left$(v, 1) = Chr$(34) And Right$(v, 1) = Chr$(34) Then v = Mid$(v, 2, Len(v) - 2)
            End If
            VbpValue = v
            Exit Function
        End If
    Next
    VbpValue = def
End Function

' Very small JSON string-field reader for the flat objects this module writes.
Private Function JsonFieldOf(ByVal chunk As String, ByVal key As String) As String
    Dim needle As String, p As Long, q As Long, s As String
    needle = Chr$(34) & key & Chr$(34) & ":" & Chr$(34)
    p = InStr(chunk, needle)
    If p = 0 Then Exit Function
    p = p + Len(needle)
    q = InStr(p, chunk, Chr$(34))
    If q = 0 Then Exit Function
    s = Mid$(chunk, p, q - p)
    s = Replace(s, "\\", "\")
    JsonFieldOf = s
End Function

Private Function AbsPath(ByVal baseFile As String, ByVal p As String) As String
    Dim t As String
    t = Replace(p, "/", "\")
    If Len(t) >= 2 Then
        If Mid$(t, 2, 1) = ":" Then
            AbsPath = t
            Exit Function
        End If
    End If
    If Left$(t, 2) = "\\" Then
        AbsPath = t
        Exit Function
    End If
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    AbsPath = fso.BuildPath(fso.GetParentFolderName(baseFile), t)
End Function

Private Function NormPath(ByVal p As String) As String
    NormPath = Replace(Replace(Trim$(p), "/", "\"), "\\", "\")
End Function

Private Function DirOf(ByVal filePath As String) As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    DirOf = fso.GetParentFolderName(filePath)
End Function

Private Sub EnsureDir(ByVal dir As String)
    ' Recursive: FileSystemObject.CreateFolder cannot create a nested path whose
    ' parent does not exist yet (it raises "path not found"), which is exactly
    ' the case for the default snapshot folder <project>\_vb6mcp_snapshots\<stamp>.
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(dir) = 0 Then Exit Sub
    If fso.FolderExists(dir) Then Exit Sub
    Dim parent As String
    parent = fso.GetParentFolderName(dir)
    If Len(parent) > 0 Then
        If Not fso.FolderExists(parent) Then EnsureDir parent
    End If
    fso.CreateFolder dir
    On Error GoTo 0
End Sub

Private Sub BackupFile(ByVal path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.CopyFile path, path & ".bak", True
    On Error GoTo 0
End Sub

Private Function Stamp() As String
    Stamp = Format$(Now, "yyyymmdd_hhnnss")
End Function

Private Function TrimRight$(ByVal s As String)
    Do While Len(s) > 0
        If Right$(s, 1) = vbLf Or Right$(s, 1) = vbCr Then
            s = Left$(s, Len(s) - 1)
        Else
            Exit Do
        End If
    Loop
    TrimRight$ = s
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

Private Function StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    If Len(s) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

Public Function ReadTextFile(ByVal path As String) As String
    Dim b() As Byte, s As String
    On Error Resume Next
    ReadAllBytesTo path, b
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReadTextFile = ""
        Exit Function
    End If
    On Error GoTo 0
    If ByteCount(b) = 0 Then
        ReadTextFile = ""
        Exit Function
    End If
    DecodeBytes b, DetectEncoding(b), s
    ReadTextFile = s
End Function

' ---- read-back evidence for every write this module performs ------------------
' Returns True ONLY when the bytes are readable back AND decode to exactly the intended text.
' It used to be a Sub over the unchecked primitive (whose Open error is swallowed), so callers
' could report a snapshot or a delivery package that was never written. VB6 allows a Function
' to be called as a statement, so existing call sites keep compiling.
Public Function WriteTextFile(ByVal path As String, ByVal s As String) As Boolean
    Dim b() As Byte, back() As Byte, got As String
    WriteTextFile = False
    EncodeText s, "ansi", b
    If Not WriteAllBytesChecked(path, b) Then Exit Function
    On Error Resume Next
    ReadAllBytesTo path, back
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    If ByteCount(back) = 0 Then Exit Function
    DecodeBytes back, DetectEncoding(back), got
    If got <> s Then Exit Function
    WriteTextFile = True
End Function

' the refusal that replaces a success shape when the read-back failed.
Private Function ShipWriteRefusal(ByVal toolName As String, ByVal target As String, ByVal reason As String, ByVal detail As String) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("tool") = toolName
    d("target") = target
    d("reason") = reason
    d("detail") = detail
    d("verdictHint") = "an artifact-producing tool may only answer success when the artifact is readable back from disk"
    ShipWriteRefusal = JsonStringify(d)
End Function
