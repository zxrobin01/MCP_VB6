Attribute VB_Name = "ProjMeta"
Option Explicit

' ---------------------------------------------------------------------------
' Text-level project metadata and application-manifest generation.
'
' The .vbp version resource (MajorVer/MinorVer/CompanyName/FileDescription/...)
' and the conditional-compilation arguments are NOT reachable through the VBIDE
' object model, so they are edited as text here. Component, reference and
' section lines are never touched; every write keeps a .bak next to the file.
' ---------------------------------------------------------------------------

' Reads every scalar key=value line of a .vbp (skipping components, references
' and [sections]) and returns {"path","exists","keys":{...},"lineEnding","sha256"}.
' ---- independent read-back evidence for a write ---------------------------
' True ONLY when the file exists AND decodes to EXACTLY the text that was to be written;
' returns the file's own sha256 and byte count. The write side of this module used to report
' the INTENDED size plus Sha256HexFile(), which on a failed write is the hash of an EMPTY
' file - a success shape for a write that never happened.
Private Function WrittenBackOk(ByVal path As String, ByVal expectedText As String, ByRef shaOut As String, ByRef bytesOut As Long) As Boolean
    Dim b() As Byte, gotText As String
    shaOut = ""
    bytesOut = 0
    WrittenBackOk = False
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(path) Then Exit Function
    On Error Resume Next
    ReadAllBytesTo path, b
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0
    bytesOut = ByteCount(b)
    If bytesOut = 0 Then Exit Function
    DecodeBytes b, DetectEncoding(b), gotText
    If gotText <> expectedText Then Exit Function
    shaOut = Sha256HexFile(path)
    WrittenBackOk = True
End Function

' the refusal that replaces a success shape when the read-back failed.
Private Function WriteNotVisibleJson(ByVal toolName As String, ByVal target As String, ByVal shaOut As String, ByVal bytesOut As Long) As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("tool") = toolName
    d("target") = target
    d("reason") = "write_not_visible"
    d("actualSha256") = shaOut
    d("actualBytes") = bytesOut
    d("detail") = "the write did NOT become visible on disk, or the bytes on disk are not the bytes that were to be written (target may be unwritable, the path invalid, or something else holds it). The success shape is deliberately NOT returned."
    d("verdictHint") = "a write tool may only answer 'written' when the bytes are readable back from the target path"
    WriteNotVisibleJson = JsonStringify(d)
End Function

Public Function MetaGetJson(ByVal vbpPath As String) As String
    Dim text As String
    text = ReadProjectText(vbpPath)
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = vbpPath
    d("exists") = (Len(text) > 0)
    d("lineEnding") = LineEndingOf(text)
    Dim keys As Object
    Set keys = CreateObject("Scripting.Dictionary")
    Dim lines() As String, i As Long
    Dim k As String, v As String
    lines = SplitLines(text)
    For i = LBound(lines) To UBound(lines)
        If ScalarKeyOf(lines(i), k, v) Then
            If Not keys.Exists(k) Then keys(k) = v
        End If
    Next
    Set d("keys") = keys
    d("sha256") = Sha256HexFile(vbpPath)
    MetaGetJson = JsonStringify(d)
End Function

' Applies {key:value} updates to a .vbp. Existing keys are rewritten in place;
' new keys are inserted before the first component / reference / [section] line.
' Returns {"path","changed","added","sha256"}; changedCsv/addedCsv mirror the lists.
Public Function MetaSetJson(ByVal vbpPath As String, ByVal updates As Object, _
                            ByRef changedCsv As String, ByRef addedCsv As String) As String
    If updates Is Nothing Then Err.Raise vbObjectError + 61, "ProjMeta", "updates is required"
    If updates.Count = 0 Then Err.Raise vbObjectError + 61, "ProjMeta", "no fields supplied to set"
    Dim text As String
    text = ReadProjectText(vbpPath)
    If Len(text) = 0 Then Err.Raise vbObjectError + 62, "ProjMeta", "project file is empty or unreadable: " & vbpPath
    Dim eol As String
    eol = LineEndingOf(text)
    Dim lines() As String
    lines = SplitLines(text)
    Dim done As Object
    Set done = CreateObject("Scripting.Dictionary")
    Dim k As Variant
    Dim mk As String, mv As String
    Dim i As Long
    changedCsv = ""
    addedCsv = ""
    For Each k In updates.Keys
        For i = LBound(lines) To UBound(lines)
            If ScalarKeyOf(lines(i), mk, mv) Then
                If StrComp(mk, CStr(k), vbTextCompare) = 0 Then
                    lines(i) = CStr(k) & "=" & CStr(updates(k))
                    done(CStr(k)) = True
                    AppendCsv changedCsv, CStr(k)
                    Exit For
                End If
            End If
        Next
    Next
    For Each k In updates.Keys
        If Not done.Exists(CStr(k)) Then
            lines = InsertMetaLine(lines, CStr(k) & "=" & CStr(updates(k)))
            AppendCsv addedCsv, CStr(k)
        End If
    Next
    BackupFile vbpPath
    Dim newText As String
    newText = Join(lines, eol)
    WriteProjectText vbpPath, newText
    ' read the .vbp back and prove the edit landed, before claiming it did.
    Dim vSha As String, vBytes As Long
    If Not WrittenBackOk(vbpPath, newText, vSha, vBytes) Then
        MetaSetJson = WriteNotVisibleJson("project_meta", vbpPath, vSha, vBytes)
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = vbpPath
    d("changed") = changedCsv
    d("added") = addedCsv
    d("sha256") = vSha
    d("bytesWritten") = vBytes
    d("verifiedOnDisk") = True
    MetaSetJson = JsonStringify(d)
End Function

' Builds an application manifest (UAC level + DPI awareness). Pure text: the
' caller writes it next to the exe (external manifest) or embeds it.
Public Function ManifestText(ByVal level As String, ByVal dpiAware As Boolean, _
                             ByVal dpiAwareness As String, ByVal uiAccess As Boolean, _
                             ByVal assemblyName As String) As String
    Dim q As String
    q = Chr$(34)
    Dim lvl As String
    lvl = LevelOf(level)
    Dim asm As String
    asm = assemblyName
    If Len(asm) = 0 Then asm = "vb6mcp.app"
    Dim s As String
    s = "<?xml version=" & q & "1.0" & q & " encoding=" & q & "UTF-8" & q & " standalone=" & q & "yes" & q & "?>" & vbCrLf
    s = s & "<assembly xmlns=" & q & "urn:schemas-microsoft-com:asm.v1" & q & " manifestVersion=" & q & "1.0" & q & ">" & vbCrLf
    s = s & "  <assemblyIdentity type=" & q & "win32" & q & " name=" & q & asm & q & " version=" & q & "1.0.0.0" & q & " processorArchitecture=" & q & "x86" & q & " />" & vbCrLf
    s = s & "  <trustInfo xmlns=" & q & "urn:schemas-microsoft-com:asm.v3" & q & ">" & vbCrLf
    s = s & "    <security>" & vbCrLf
    s = s & "      <requestedPrivileges>" & vbCrLf
    s = s & "        <requestedExecutionLevel level=" & q & lvl & q & " uiAccess=" & q & BoolText(uiAccess) & q & " />" & vbCrLf
    s = s & "      </requestedPrivileges>" & vbCrLf
    s = s & "    </security>" & vbCrLf
    s = s & "  </trustInfo>" & vbCrLf
    s = s & "  <application xmlns=" & q & "urn:schemas-microsoft-com:asm.v3" & q & ">" & vbCrLf
    s = s & "    <windowsSettings>" & vbCrLf
    s = s & "      <dpiAware xmlns=" & q & "http://schemas.microsoft.com/SMI/2005/WindowsSettings" & q & ">" & BoolText(dpiAware) & "</dpiAware>" & vbCrLf
    Dim da As String
    da = DpiAwarenessOf(dpiAwareness)
    If Len(da) > 0 Then
        s = s & "      <dpiAwareness xmlns=" & q & "http://schemas.microsoft.com/SMI/2016/WindowsSettings" & q & ">" & da & "</dpiAwareness>" & vbCrLf
    End If
    s = s & "    </windowsSettings>" & vbCrLf
    s = s & "  </application>" & vbCrLf
    s = s & "</assembly>" & vbCrLf
    ManifestText = s
End Function

' Writes <exePath>.manifest and returns {"path","bytes","level","dpiAwareness","sha256"}.
Public Function ManifestWriteJson(ByVal exePath As String, ByVal level As String, _
                                  ByVal dpiAware As Boolean, ByVal dpiAwareness As String, _
                                  ByVal uiAccess As Boolean, ByVal assemblyName As String) As String
    If Len(exePath) = 0 Then Err.Raise vbObjectError + 63, "ProjMeta", "path is required (the .exe or .dll the manifest belongs to)"
    Dim s As String
    s = ManifestText(level, dpiAware, dpiAwareness, uiAccess, assemblyName)
    Dim manPath As String
    manPath = exePath & ".manifest"
    BackupFile manPath
    WriteProjectText manPath, s
    ' same read-back assertion as MetaSetJson.
    Dim mSha As String, mBytes As Long
    If Not WrittenBackOk(manPath, s, mSha, mBytes) Then
        ManifestWriteJson = WriteNotVisibleJson("app_manifest", manPath, mSha, mBytes)
        Exit Function
    End If
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = manPath
    d("bytes") = mBytes
    d("bytesWritten") = mBytes
    d("verifiedOnDisk") = True
    d("level") = LevelOf(level)
    d("dpiAware") = dpiAware
    d("dpiAwareness") = DpiAwarenessOf(dpiAwareness)
    d("uiAccess") = uiAccess
    d("sha256") = mSha
    d("note") = "External manifest: it must sit next to the exe with the same name plus .manifest. Redeploy the exe afterwards."
    ManifestWriteJson = JsonStringify(d)
End Function

' Reads an existing <exePath>.manifest, or reports that there is none.
Public Function ManifestReadJson(ByVal exePath As String) As String
    Dim manPath As String
    manPath = exePath & ".manifest"
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("path") = manPath
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(manPath) Then
        d("exists") = False
    Else
        d("exists") = True
        d("text") = ReadProjectText(manPath)
        d("sha256") = Sha256HexFile(manPath)
    End If
    ManifestReadJson = JsonStringify(d)
End Function

' --------------------------------------------------------------- helpers

Private Function LevelOf(ByVal level As String) As String
    Dim t As String
    t = LCase$(Trim$(level))
    Select Case t
        Case "", "asinvoker"
            LevelOf = "asInvoker"
        Case "highestavailable", "highest"
            LevelOf = "highestAvailable"
        Case "requireadministrator", "admin", "elevated"
            LevelOf = "requireAdministrator"
        Case Else
            Err.Raise vbObjectError + 64, "ProjMeta", "Invalid level: " & level & " (expected asInvoker/highestAvailable/requireAdministrator)"
    End Select
End Function

Private Function DpiAwarenessOf(ByVal v As String) As String
    Dim t As String
    t = LCase$(Trim$(v))
    Select Case t
        Case "permonitorv2", "permonitor2"
            DpiAwarenessOf = "PerMonitorV2"
        Case "permonitor"
            DpiAwarenessOf = "PerMonitor"
        Case "system"
            DpiAwarenessOf = "System"
        Case Else
            DpiAwarenessOf = ""
    End Select
End Function

Private Function BoolText(ByVal b As Boolean) As String
    If b Then BoolText = "true" Else BoolText = "false"
End Function

Private Sub AppendCsv(ByRef csv As String, ByVal item As String)
    If Len(csv) > 0 Then csv = csv & ","
    csv = csv & item
End Sub

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

' A scalar key=value line: not a [section], not a component/reference/object line.
Private Function ScalarKeyOf(ByVal ln As String, ByRef k As String, ByRef v As String) As Boolean
    Dim t As String, eq As Long
    k = ""
    v = ""
    t = Trim$(ln)
    If Len(t) = 0 Then Exit Function
    If Left$(t, 1) = "[" Then Exit Function
    If IsComponentLine(t) Then Exit Function
    eq = InStr(t, "=")
    If eq <= 1 Then Exit Function
    k = Trim$(Left$(t, eq - 1))
    v = Trim$(Mid$(t, eq + 1))
    If Len(k) = 0 Then Exit Function
    ScalarKeyOf = True
End Function

Private Function IsComponentLine(ByVal t As String) As Boolean
    If StartsWith(t, "Module=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "Class=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "Form=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "UserControl=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "RelatedDoc=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "ResFile=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "Reference=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "Object=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "Designer=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "UserDocument=") Then IsComponentLine = True: Exit Function
    If StartsWith(t, "PropertyPage=") Then IsComponentLine = True: Exit Function
End Function

Private Function StartsWith(ByVal s As String, ByVal prefix As String) As Boolean
    If Len(s) < Len(prefix) Then Exit Function
    StartsWith = (StrComp(Left$(s, Len(prefix)), prefix, vbTextCompare) = 0)
End Function

Private Function InsertMetaLine(ByRef lines() As String, ByVal newLine As String) As String()
    Dim n As Long, i As Long, j As Long, at As Long
    n = UBound(lines) - LBound(lines) + 1
    at = InsertIndex(lines)
    Dim outp() As String
    ReDim outp(0 To n)
    j = 0
    For i = LBound(lines) To UBound(lines)
        If i = at Then
            outp(j) = newLine
            j = j + 1
        End If
        outp(j) = lines(i)
        j = j + 1
    Next
    If at > UBound(lines) Then outp(j) = newLine
    InsertMetaLine = outp
End Function

Private Function InsertIndex(ByRef lines() As String) As Long
    Dim i As Long, t As String
    For i = LBound(lines) To UBound(lines)
        t = Trim$(lines(i))
        If Len(t) > 0 Then
            If Left$(t, 1) = "[" Then
                InsertIndex = i
                Exit Function
            End If
            If IsComponentLine(t) Then
                InsertIndex = i
                Exit Function
            End If
        End If
    Next
    InsertIndex = UBound(lines) + 1
End Function

Private Function ReadProjectText(ByVal path As String) As String
    Dim b() As Byte
    On Error Resume Next
    ReadAllBytesTo path, b
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ReadProjectText = ""
        Exit Function
    End If
    On Error GoTo 0
    If ByteCount(b) = 0 Then
        ReadProjectText = ""
        Exit Function
    End If
    Dim enc As String, s As String
    enc = DetectEncoding(b)
    DecodeBytes b, enc, s
    ReadProjectText = s
End Function

Private Sub WriteProjectText(ByVal path As String, ByVal s As String)
    Dim b() As Byte
    EncodeText s, "ansi", b
    ' the checked writer - the caller's WrittenBackOk() assertion then compares
    ' the decoded TEXT, this compares the BYTES. Two independent read-backs.
    WriteAllBytesChecked path, b
End Sub

Private Sub BackupFile(ByVal path As String)
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(path) Then fso.CopyFile path, path & ".bak", True
    On Error GoTo 0
End Sub
