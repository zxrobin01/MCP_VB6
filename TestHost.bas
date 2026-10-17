Attribute VB_Name = "TestHost"
Option Explicit

Public Function Q() As String
    Q = Chr$(34)
End Function

' remaining: both writers now READ THE BYTES BACK and return the verdict, so
' test_host{action=generate} can refuse instead of reporting a host it never wrote.
Public Function WriteAnsi(ByVal path As String, ByVal text As String) As Boolean
    Dim b() As Byte
    EncodeText text, "ansi", b
    WriteAnsi = WriteAllBytesChecked(path, b)
End Function

Public Function GenerateAssertionsModule(ByVal templatePath As String, ByVal outPath As String) As Boolean
    Dim b() As Byte
    ReadAllBytesTo templatePath, b
    If ByteCount(b) <= 0 Then Exit Function
    GenerateAssertionsModule = WriteAllBytesChecked(outPath, b)
End Function

Public Function GenerateTestHostVbp(ByVal vbpPath As String, ByVal moduleUnderTestPath As String, ByVal assertionsPath As String, ByVal runnerPath As String, ByVal exeName As String) As Boolean
    ' Kept for compatibility: the original implementation hard-coded the
    ' component name "PlanFactLite". Delegates to the parameterized form.
    GenerateTestHostVbp = GenerateTestHostVbpNamed(vbpPath, "PlanFactLite", moduleUnderTestPath, assertionsPath, runnerPath, exeName)
End Function

Public Function GenerateTestHostVbpNamed(ByVal vbpPath As String, ByVal moduleName As String, ByVal moduleUnderTestPath As String, ByVal assertionsPath As String, ByVal runnerPath As String, ByVal exeName As String) As Boolean
    Dim s As String
    s = "Type=Exe" & vbCrLf
    ' The reference path is written as the VB6 macro $(WinDir) on purpose: this file is GENERATED
    ' on the machine that runs the test host, so a drive letter (or any absolute path) would make
    ' the generated .vbp describe ONE host. The macro expands on whichever machine loads it.
    s = s & "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#$(WinDir)\SysWOW64\stdole2.tlb#OLE Automation" & vbCrLf
    If Len(moduleName) = 0 Then moduleName = "PlanFactLite"
    s = s & "Module=" & moduleName & "; " & moduleUnderTestPath & vbCrLf
    s = s & "Module=modTest; " & assertionsPath & vbCrLf
    s = s & "Module=modTestRunner; " & runnerPath & vbCrLf
    s = s & "Startup=" & Q() & "Sub Main" & Q() & vbCrLf
    s = s & "HelpFile=" & Q() & Q() & vbCrLf
    s = s & "Title=" & Q() & "vb6testhost" & Q() & vbCrLf
    s = s & "ExeName32=" & Q() & exeName & Q() & vbCrLf
    s = s & "Name=" & Q() & "vb6testhost" & Q() & vbCrLf
    GenerateTestHostVbpNamed = WriteAnsi(vbpPath, s)
End Function

Public Function CountResults(ByVal resultPath As String, ByRef passN As Long, ByRef failN As Long) As Long
    Dim s As String, lines() As String, i As Long, n As Long
    s = ReadLogAnsi(resultPath)
    lines = Split(Replace(Replace(s, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    passN = 0
    failN = 0
    n = 0
    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            n = n + 1
            If InStr(lines(i), Q() & "status" & Q() & ":" & Q() & "PASS" & Q()) > 0 Then passN = passN + 1
            If InStr(lines(i), Q() & "status" & Q() & ":" & Q() & "FAIL" & Q()) > 0 Then failN = failN + 1
        End If
    Next
    CountResults = n
End Function
