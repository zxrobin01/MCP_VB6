Attribute VB_Name = "Encoding"
Option Explicit

Private Const CP_UTF8 As Long = 65001
Private Const CP_ACP As Long = 0
Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal cp As Long, ByVal flags As Long, ByVal lpMB As Long, ByVal cbMB As Long, ByVal lpWC As Long, ByVal cchWC As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal cp As Long, ByVal flags As Long, ByVal lpWC As Long, ByVal cchWC As Long, ByVal lpMB As Long, ByVal cbMB As Long, ByVal lpDef As Long, ByVal lpUsed As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDst As Long, ByVal pSrc As Long, ByVal n As Long)

Public Function ByteCount(ByRef b() As Byte) As Long
    On Error Resume Next
    ByteCount = 0
    ByteCount = UBound(b) - LBound(b) + 1
    On Error GoTo 0
End Function

Public Function DetectEncoding(ByRef b() As Byte) As String
    Dim n As Long, p As Long, off As Long, cch As Long
    n = ByteCount(b)
    If n <= 0 Then
        DetectEncoding = "empty"
        Exit Function
    End If
    p = LBound(b)
    If n >= 2 Then
        If b(p) = &HFF And b(p + 1) = &HFE Then
            DetectEncoding = "utf-16le"
            Exit Function
        End If
        If b(p) = &HFE And b(p + 1) = &HFF Then
            DetectEncoding = "utf-16be"
            Exit Function
        End If
    End If
    off = 0
    If n >= 3 Then
        If b(p) = &HEF And b(p + 1) = &HBB And b(p + 2) = &HBF Then off = 3
    End If
    cch = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, VarPtr(b(p + off)), n - off, 0, 0)
    If cch > 0 Then
        DetectEncoding = "utf-8"
    Else
        DetectEncoding = "ansi"
    End If
End Function

Public Sub DecodeBytes(ByRef b() As Byte, ByVal enc As String, ByRef out As String)
    Dim n As Long, p As Long, off As Long, cch As Long, units As Long, cp As Long
    out = ""
    n = ByteCount(b)
    If n <= 0 Then Exit Sub
    p = LBound(b)
    If enc = "utf-16le" Then
        off = 0
        If n >= 2 Then
            If b(p) = &HFF And b(p + 1) = &HFE Then off = 2
        End If
        units = (n - off) \ 2
        If units <= 0 Then Exit Sub
        out = String$(units, vbNullChar)
        CopyMemory StrPtr(out), VarPtr(b(p + off)), units * 2
        Exit Sub
    End If
    off = 0
    If enc = "utf-8" Then
        If n >= 3 Then
            If b(p) = &HEF And b(p + 1) = &HBB And b(p + 2) = &HBF Then off = 3
        End If
        cp = CP_UTF8
    Else
        cp = CP_ACP
    End If
    cch = MultiByteToWideChar(cp, 0, VarPtr(b(p + off)), n - off, 0, 0)
    If cch <= 0 Then Exit Sub
    out = String$(cch, vbNullChar)
    MultiByteToWideChar cp, 0, VarPtr(b(p + off)), n - off, StrPtr(out), cch
End Sub

Public Sub EncodeText(ByVal s As String, ByVal enc As String, ByRef out() As Byte)
    Dim n As Long, cb As Long, cp As Long
    If Len(s) = 0 Then Exit Sub
    If enc = "utf-16le" Then
        n = Len(s)
        ReDim out(0 To n * 2 + 1)
        out(0) = &HFF
        out(1) = &HFE
        CopyMemory VarPtr(out(2)), StrPtr(s), n * 2
        Exit Sub
    End If
    If enc = "utf-8" Then cp = CP_UTF8 Else cp = CP_ACP
    cb = WideCharToMultiByte(cp, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    If cb <= 0 Then Exit Sub
    ReDim out(0 To cb - 1)
    WideCharToMultiByte cp, 0, StrPtr(s), Len(s), VarPtr(out(0)), cb, 0, 0
End Sub

Public Sub ReadAllBytesTo(ByVal path As String, ByRef out() As Byte)
    Dim fnum As Integer, n As Long
    fnum = FreeFile
    On Error GoTo fail
    Open path For Binary Access Read As #fnum
    n = LOF(fnum)
    If n > 0 Then
        ReDim out(0 To n - 1)
        Get #fnum, , out
    End If
    Close #fnum
    Exit Sub
fail:
    On Error Resume Next
    Close #fnum
End Sub

Public Sub WriteAllBytesTo(ByVal path As String, ByRef b() As Byte)
    Dim fnum As Integer, n As Long
    n = ByteCount(b)
    ' VB6 Open ... For Binary Access Write does NOT truncate an existing file:
    ' a shorter write leaves the old tail behind. Always delete first.
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    On Error GoTo fail
    fnum = FreeFile
    Open path For Binary Access Write As #fnum
    If n > 0 Then Put #fnum, , b
    Close #fnum
    Exit Sub
fail:
    On Error Resume Next
    Close #fnum
End Sub

Public Function Sha256HexFile(ByVal path As String) As String
    Dim b() As Byte
    ReadAllBytesTo path, b
    Sha256HexFile = Sha256HexBytes(b)
End Function
