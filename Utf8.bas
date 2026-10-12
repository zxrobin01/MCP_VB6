Attribute VB_Name = "Utf8"
Option Explicit

' UTF-8 / ANSI / UTF-16 conversion helpers (kernel32 based, no ADO dependency).
' A "byte-per-char string" is a VB6 string where each character holds one raw
' byte (0..255); it is used as the wire format between Winsock and the parser.

Private Const CP_UTF8 As Long = 65001
Private Const CP_ACP As Long = 0
Private Const MB_ERR_INVALID_CHARS As Long = &H8

Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal pDest As Long, ByVal pSrc As Long, ByVal ByteLen As Long)

Public Function Utf8Encode(ByVal s As String) As Byte()
    Dim n As Long
    n = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        WideCharToMultiByte CP_UTF8, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    End If
    Utf8Encode = b
End Function

Public Function AnsiEncode(ByVal s As String) As Byte()
    Dim n As Long
    n = WideCharToMultiByte(CP_ACP, 0, StrPtr(s), Len(s), 0, 0, 0, 0)
    Dim b() As Byte
    If n > 0 Then
        ReDim b(0 To n - 1)
        WideCharToMultiByte CP_ACP, 0, StrPtr(s), Len(s), VarPtr(b(0)), n, 0, 0
    End If
    AnsiEncode = b
End Function

Public Function Utf16Encode(ByVal s As String) As Byte()
    ' little-endian UTF-16 with BOM
    Dim n As Long
    n = Len(s)
    Dim b() As Byte
    ReDim b(0 To n * 2 + 1)
    b(0) = &HFF
    b(1) = &HFE
    If n > 0 Then
        CopyMemory VarPtr(b(2)), StrPtr(s), n * 2
    End If
    Utf16Encode = b
End Function

Public Function Utf8DecodeBytes(b() As Byte) As String
    Dim n As Long
    On Error Resume Next
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Or n <= 0 Then Exit Function
    On Error GoTo 0
    Dim cch As Long
    cch = MultiByteToWideChar(CP_UTF8, 0, VarPtr(b(LBound(b))), n, 0, 0)
    If cch <= 0 Then Exit Function
    Dim s As String
    s = String$(cch, vbNullChar)
    MultiByteToWideChar CP_UTF8, 0, VarPtr(b(LBound(b))), n, StrPtr(s), cch
    Utf8DecodeBytes = s
End Function

Public Function AnsiDecodeBytes(b() As Byte) As String
    Dim n As Long
    On Error Resume Next
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Or n <= 0 Then Exit Function
    On Error GoTo 0
    Dim cch As Long
    cch = MultiByteToWideChar(CP_ACP, 0, VarPtr(b(LBound(b))), n, 0, 0)
    If cch <= 0 Then Exit Function
    Dim s As String
    s = String$(cch, vbNullChar)
    MultiByteToWideChar CP_ACP, 0, VarPtr(b(LBound(b))), n, StrPtr(s), cch
    AnsiDecodeBytes = s
End Function

Public Function TryUtf8Decode(b() As Byte, ByRef outText As String) As Boolean
    ' returns False when the byte sequence is not valid UTF-8
    Dim n As Long
    On Error Resume Next
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Or n <= 0 Then
        outText = ""
        TryUtf8Decode = True
        Exit Function
    End If
    On Error GoTo 0
    Dim cch As Long
    cch = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, VarPtr(b(LBound(b))), n, 0, 0)
    If cch <= 0 Then Exit Function
    Dim s As String
    s = String$(cch, vbNullChar)
    MultiByteToWideChar CP_UTF8, MB_ERR_INVALID_CHARS, VarPtr(b(LBound(b))), n, StrPtr(s), cch
    outText = s
    TryUtf8Decode = True
End Function

Public Function Utf16DecodeBytes(b() As Byte) As String
    Dim n As Long
    On Error Resume Next
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Or n < 2 Then Exit Function
    On Error GoTo 0
    Dim s As String
    s = String$((n \ 2), vbNullChar)
    CopyMemory StrPtr(s), VarPtr(b(LBound(b))), (n \ 2) * 2
    If Left$(s, 1) = ChrW$(&HFEFF) Then s = Mid$(s, 2)
    Utf16DecodeBytes = s
End Function

Public Function Utf8DecodeStr(ByVal s As String) As String
    ' decode a byte-per-char string
    Utf8DecodeStr = Utf8DecodeBytes(StrToBytes(s))
End Function

Public Function StrToBytes(ByVal s As String) As Byte()
    ' byte-per-char string -> byte array (identity mapping, code-page safe)
    Dim n As Long
    n = Len(s)
    Dim b() As Byte
    If n <= 0 Then
        StrToBytes = b
        Exit Function
    End If
    ReDim b(0 To n - 1)
    Dim i As Long
    For i = 1 To n
        b(i - 1) = CByte(AscW(Mid$(s, i, 1)) And &HFF&)
    Next
    StrToBytes = b
End Function

Public Function BytesToStr(b() As Byte) As String
    ' byte array -> byte-per-char string (identity mapping, code-page safe)
    Dim n As Long
    n = UBound(b) - LBound(b) + 1
    If n <= 0 Then Exit Function
    Dim s As String
    s = String$(n, vbNullChar)
    Dim i As Long
    For i = 0 To n - 1
        Mid$(s, i + 1, 1) = ChrW$(b(LBound(b) + i))
    Next
    BytesToStr = s
End Function
