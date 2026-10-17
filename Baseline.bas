Attribute VB_Name = "Baseline"
Option Explicit

Public Function BaseName(ByVal path As String) As String
    Dim p As Long
    p = InStrRev(path, Chr$(92))
    If p = 0 Then p = InStrRev(path, "/")
    If p = 0 Then
        BaseName = path
    Else
        BaseName = Mid$(path, p + 1)
    End If
End Function

Public Function SignatureOf(ByVal file As String, ByVal reportedLine As Long, ByVal key As String) As String
    SignatureOf = LCase$(BaseName(file)) & "|" & CStr(reportedLine) & "|" & LCase$(key)
End Function

Public Function SignatureOfBuild(ByRef br As TBuildResult) As String
    SignatureOfBuild = SignatureOf(br.file, br.reportedLine, br.key)
End Function

Public Function LoadBaselineText(ByVal path As String) As String
    LoadBaselineText = ReadLogAnsi(path)
End Function

Public Function SaveBaselineText(ByVal path As String, ByVal text As String) As Boolean
    Dim b() As Byte
    On Error GoTo fail
    ' remaining: report success only when the bytes are READABLE BACK.
    EncodeText text, "ansi", b
    SaveBaselineText = WriteAllBytesChecked(path, b)
    Exit Function
fail:
    SaveBaselineText = False
End Function

Public Function BaselineCount(ByVal baselineLines As String) As Long
    Dim bl() As String, i As Long, n As Long
    bl = Split(Replace(Replace(baselineLines, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(bl)
        If Len(Trim$(bl(i))) > 0 Then n = n + 1
    Next
    BaselineCount = n
End Function

Public Function BaselineHas(ByVal baselineLines As String, ByVal sig As String) As Boolean
    Dim bl() As String, i As Long
    bl = Split(Replace(Replace(baselineLines, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    For i = 0 To UBound(bl)
        If StrComp(Trim$(bl(i)), sig, vbTextCompare) = 0 Then
            BaselineHas = True
            Exit Function
        End If
    Next
End Function

Public Function BuildDiffJson(ByVal currentOk As Boolean, ByVal curSig As String, ByVal baselineLines As String) As String
    Dim blCount As Long, newE As Long, fixedE As Long, unchanged As Long, inBase As Boolean, verdict As String
    blCount = BaselineCount(baselineLines)
    inBase = False
    If blCount > 0 And Len(curSig) > 0 Then inBase = BaselineHas(baselineLines, curSig)
    If currentOk Then
        fixedE = blCount
        verdict = "OK"
        If blCount > 0 Then verdict = "FIXED"
    Else
        If inBase Then
            unchanged = 1
            verdict = "UNCHANGED"
        Else
            newE = 1
            verdict = "NEW-ERROR"
        End If
    End If
    BuildDiffJson = "{" & Jq("newErrors") & ":" & CStr(newE) & "," & Jq("fixedErrors") & ":" & CStr(fixedE) & "," & Jq("unchanged") & ":" & CStr(unchanged) & "," & Jq("baselineCount") & ":" & CStr(blCount) & "," & Jq("verdict") & ":" & Js(verdict) & "," & Jq("currentSignature") & ":" & Js(curSig) & "}"
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
