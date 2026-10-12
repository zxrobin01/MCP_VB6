Attribute VB_Name = "Fixture"
Option Explicit
' C11 - deterministic sample (fixture) record + compare.
' Storage is raw ANSI bytes so a round trip is byte exact; the compare is
' line based (set membership plus first differing line).

Public Function FixSave(ByVal path As String, ByVal text As String) As Long
    Dim fn As Integer, b() As Byte
    On Error GoTo Bad
    b = StrConv(text, vbFromUnicode)
    ' Open ... For Binary Access Write does NOT truncate: kill first.
    On Error Resume Next
    Kill path
    On Error GoTo Bad
    fn = FreeFile
    Open path For Binary Access Write As #fn
    If Len(text) > 0 Then Put #fn, , b
    Close #fn
    FixSave = Len(text)
    Exit Function
Bad:
    On Error Resume Next
    Close #fn
    FixSave = -1
End Function

Public Function FixLoad(ByVal path As String) As String
    Dim fn As Integer, b() As Byte, n As Long
    On Error GoTo Bad
    fn = FreeFile
    Open path For Binary Access Read As #fn
    n = LOF(fn)
    If n > 0 Then
        ReDim b(0 To n - 1)
        Get #fn, , b
    End If
    Close #fn
    If n > 0 Then FixLoad = StrConv(b, vbUnicode) Else FixLoad = ""
    Exit Function
Bad:
    On Error Resume Next
    Close #fn
    FixLoad = ""
End Function

Public Function FixRecordJson(ByVal path As String, ByVal text As String) As String
    Dim n As Long
    n = FixSave(path, text)
    If n < 0 Then
        FixRecordJson = "{" & Qs("ok") & ":false," & Qs("error") & ":" & Js("cannot write fixture") & "}"
    Else
        FixRecordJson = "{" & Qs("ok") & ":true," & Qs("bytes") & ":" & CStr(n) & "," & Qs("path") & ":" & Js(path) & "}"
    End If
End Function

Public Function FixCompareJson(ByVal path As String, ByVal text As String) As String
    Dim stored As String
    stored = FixLoad(path)
    Dim nStored As Long, nCur As Long
    Dim a() As String, b() As String
    a = Split(Norm(stored), vbLf)
    b = Split(Norm(text), vbLf)
    nStored = UBound(a) + 1
    nCur = UBound(b) + 1
    Dim i As Long, firstDiff As Long, lim As Long
    firstDiff = 0
    lim = nStored
    If nCur < lim Then lim = nCur
    For i = 0 To lim - 1
        If a(i) <> b(i) Then firstDiff = i + 1 : Exit For
    Next
    If firstDiff = 0 And nStored <> nCur Then firstDiff = lim + 1
    Dim da As Object, db As Object
    Set da = CreateObject("Scripting.Dictionary")
    Set db = CreateObject("Scripting.Dictionary")
    For i = 0 To nStored - 1
        da(a(i)) = 1
    Next
    For i = 0 To nCur - 1
        db(b(i)) = 1
    Next
    Dim removed As Long, added As Long
    For i = 0 To nStored - 1
        If Not db.Exists(a(i)) Then removed = removed + 1
    Next
    For i = 0 To nCur - 1
        If Not da.Exists(b(i)) Then added = added + 1
    Next
    Dim same As Boolean
    same = (removed = 0 And added = 0 And firstDiff = 0)
    FixCompareJson = "{" & Qs("exists") & ":" & LCase$(CStr(Len(stored) > 0))
    FixCompareJson = FixCompareJson & "," & Qs("same") & ":" & LCase$(CStr(same))
    FixCompareJson = FixCompareJson & "," & Qs("storedLines") & ":" & CStr(nStored)
    FixCompareJson = FixCompareJson & "," & Qs("currentLines") & ":" & CStr(nCur)
    FixCompareJson = FixCompareJson & "," & Qs("added") & ":" & CStr(added)
    FixCompareJson = FixCompareJson & "," & Qs("removed") & ":" & CStr(removed)
    FixCompareJson = FixCompareJson & "," & Qs("firstDiffLine") & ":" & CStr(firstDiff) & "}"
End Function

Private Function Norm(ByVal s As String) As String
    Dim t As String
    t = Replace(s, vbCrLf, vbLf)
    t = Replace(t, vbCr, vbLf)
    Norm = t
End Function

Private Function Qs(ByVal s As String) As String
    Qs = """" & s & """"
End Function

Private Function Js(ByVal s As String) As String
    Dim t As String, b As String
    b = Chr$(92)
    t = Replace(s, b, b & b)
    t = Replace(t, """", b & """")
    Js = """" & t & """"
End Function
