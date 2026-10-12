Attribute VB_Name = "TestHost"
Option Explicit

Public Function Q() As String
    Q = Chr$(34)
End Function

Public Sub WriteAnsi(ByVal path As String, ByVal text As String)
    Dim b() As Byte
    EncodeText text, "ansi", b
    WriteAllBytesTo path, b
End Sub

Public Sub GenerateAssertionsModule(ByVal templatePath As String, ByVal outPath As String)
    Dim b() As Byte
    ReadAllBytesTo templatePath, b
    WriteAllBytesTo outPath, b
End Sub

Public Sub GenerateTestHostVbp(ByVal vbpPath As String, ByVal moduleUnderTestPath As String, ByVal assertionsPath As String, ByVal runnerPath As String, ByVal exeName As String)
    ' Kept for compatibility: the original implementation hard-coded the
    ' component name "PlanFactLite". Delegates to the parameterized form.
    GenerateTestHostVbpNamed vbpPath, "PlanFactLite", moduleUnderTestPath, assertionsPath, runnerPath, exeName
End Sub

Public Sub GenerateTestHostVbpNamed(ByVal vbpPath As String, ByVal moduleName As String, ByVal moduleUnderTestPath As String, ByVal assertionsPath As String, ByVal runnerPath As String, ByVal exeName As String)
    Dim s As String
    s = "Type=Exe" & vbCrLf
    s = s & "Reference=*\G{00020430-0000-0000-C000-000000000046}#2.0#0#B:\Windows\SysWOW64\stdole2.tlb#OLE Automation" & vbCrLf
    If Len(moduleName) = 0 Then moduleName = "PlanFactLite"
    s = s & "Module=" & moduleName & "; " & moduleUnderTestPath & vbCrLf
    s = s & "Module=modTest; " & assertionsPath & vbCrLf
    s = s & "Module=modTestRunner; " & runnerPath & vbCrLf
    s = s & "Startup=" & Q() & "Sub Main" & Q() & vbCrLf
    s = s & "HelpFile=" & Q() & Q() & vbCrLf
    s = s & "Title=" & Q() & "vb6testhost" & Q() & vbCrLf
    s = s & "ExeName32=" & Q() & exeName & Q() & vbCrLf
    s = s & "Name=" & Q() & "vb6testhost" & Q() & vbCrLf
    WriteAnsi vbpPath, s
End Sub

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
