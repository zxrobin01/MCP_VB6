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
        ' STRICT utf-8.  The old call here used flags=0, i.e. a REPLACEMENT
        ' decode: every invalid byte sequence silently became U+FFFD and then
        ' travelled on into a code pane.  A buffer that is not valid utf-8 is decoded
        ' as ACP (the codepage VB6 itself uses for these files) instead - the write
        ' path must never be able to produce U+FFFD out of bytes that were not
        ' literally U+FFFD.
        cch = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, VarPtr(b(p + off)), n - off, 0, 0)
        If cch > 0 Then
            cp = CP_UTF8
        Else
            cp = CP_ACP
        End If
    Else
        cp = CP_ACP
    End If
    cch = MultiByteToWideChar(cp, 0, VarPtr(b(p + off)), n - off, 0, 0)
    If cch <= 0 Then Exit Sub
    out = String$(cch, vbNullChar)
    MultiByteToWideChar cp, 0, VarPtr(b(p + off)), n - off, StrPtr(out), cch
End Sub

' R-1021-04 (item 4): THE decode rule for build text - the compiler log read by
' Build.ReadLogAny and the source line quoted by firstError. A buffer is decoded as
' UTF-8 ONLY when it is VALID UTF-8 (strict MultiByteToWideChar, no replacement);
' everything else is decoded with the system ANSI codepage (936 on a Chinese
' Windows, ACP elsewhere). This is what stops a UTF-8 message being decoded as GBK
' and reaching a receipt as mojibake. Returns False + why when nothing decodes, so
' a caller can report unreadable instead of printing garbage.
Public Function DecodeUtf8OrAnsi(ByRef b() As Byte, ByRef out As String, ByRef encUsed As String, ByRef why As String) As Boolean
    Dim n As Long, p As Long, cch As Long
    out = ""
    encUsed = ""
    why = ""
    DecodeUtf8OrAnsi = False
    n = ByteCount(b)
    If n <= 0 Then
        why = "empty"
        Exit Function
    End If
    p = LBound(b)
    encUsed = DetectEncoding(b)
    If encUsed = "empty" Then encUsed = "ansi"
    ' utf-16be has no decoder here; ANSI at least keeps the byte order explicit
    If encUsed = "utf-16be" Then encUsed = "ansi"
    If encUsed = "utf-8" Then
        ' STRICT re-check on the WHOLE buffer: a buffer that is not valid UTF-8 must
        ' never take this branch, whatever DetectEncoding concluded.
        cch = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, VarPtr(b(p)), n, 0, 0)
        If cch <= 0 Then encUsed = "ansi"
    End If
    DecodeBytes b, encUsed, out
    If Len(out) = 0 Then
        why = "decode_empty:" & encUsed
        Exit Function
    End If
    DecodeUtf8OrAnsi = True
End Function

Public Function BytesEqual(ByRef a() As Byte, ByRef b() As Byte, ByVal aOff As Long, ByVal bOff As Long, ByVal n As Long) As Boolean
    ' windowed byte compare between two Byte arrays (no StrConv detour).
    Dim i As Long
    BytesEqual = False
    If n < 0 Then Exit Function
    On Error GoTo bad
    If n = 0 Then
        BytesEqual = True
        Exit Function
    End If
    If aOff < LBound(a) Or bOff < LBound(b) Then Exit Function
    If aOff + n - 1 > UBound(a) Or bOff + n - 1 > UBound(b) Then Exit Function
    For i = 0 To n - 1
        If a(aOff + i) <> b(bOff + i) Then Exit Function
    Next i
    BytesEqual = True
    Exit Function
bad:
    BytesEqual = False
End Function

' the write path's encoding gate.  Detect -> decode -> RE-ENCODE with the
' SAME encoding -> compare byte for byte.  A decode that lost ANY information cannot
' round-trip, so U+FFFD can no longer be smuggled into a code pane: the caller gets
' False + why and refuses structurally (add_code{file=} -> unreadable_source_encoding).
Public Function DecodeVerified(ByRef b() As Byte, ByRef encUsed As String, ByRef out As String, ByRef why As String) As Boolean
    Dim back() As Byte, n As Long, p As Long, off As Long, sk As Long
    out = "": encUsed = "": why = ""
    DecodeVerified = False
    On Error GoTo fail
    n = ByteCount(b)
    If n <= 0 Then
        why = "empty"
        Exit Function
    End If
    p = LBound(b)
    encUsed = DetectEncoding(b)
    If encUsed = "empty" Then encUsed = "ansi"
    If encUsed = "utf-16be" Then encUsed = "ansi"
    off = 0
    sk = 0
    If encUsed = "utf-8" Then
        If n >= 3 Then
            If b(p) = &HEF And b(p + 1) = &HBB And b(p + 2) = &HBF Then off = 3
        End If
    ElseIf encUsed = "utf-16le" Then
        If n >= 2 Then
            If b(p) = &HFF And b(p + 1) = &HFE Then off = 2
        End If
        sk = 2    ' EncodeText re-emits the utf-16le BOM it consumed
    End If
    ' R-1021-04 (item 4): the write/read-back path decodes through the SAME rule as
    ' the log reader, and refuses (rather than silently disagreeing) if the rule and
    ' the encoding it pre-computed diverge.
    Dim decEnc As String
    Dim decWhy As String
    If Not DecodeUtf8OrAnsi(b, out, decEnc, decWhy) Then
        why = "decode_failed:" & decWhy
        Exit Function
    End If
    If decEnc <> encUsed Then
        why = "decode_rule_mismatch:" & encUsed & "<->" & decEnc
        Exit Function
    End If
    If Len(out) = 0 Then
        Exit Function
    End If
    EncodeText out, encUsed, back
    If ByteCount(back) - sk <> n - off Then
        why = "roundtrip_len:" & encUsed & ":" & CStr(ByteCount(back)) & "<->" & CStr(n - off)
        Exit Function
    End If
    If Not BytesEqual(b, back, p + off, LBound(back) + sk, n - off) Then
        why = "roundtrip_bytes:" & encUsed
        Exit Function
    End If
    DecodeVerified = True
    Exit Function
fail:
    why = "decode_error:" & CStr(Err.Number)
End Function

' encode -> decode -> compare.  A CP_ACP encode does NOT fail on a character
' the codepage cannot carry: it SILENTLY substitutes the default character (measured:
' an emoji became 3F 3F).  The write paths therefore PROVE the round trip instead of
' trusting the encoder.  On False, firstDiffAt (1-based, 0 = length only) locates it.
Public Function TextEncodableEx(ByVal s As String, ByVal enc As String, ByRef back As String, ByRef firstDiffAt As Long) As Boolean
    Dim b() As Byte, n As Long, i As Long, lo As Long, hi As Long, blk As Long
    back = "": firstDiffAt = 0
    TextEncodableEx = False
    If Len(s) = 0 Then
        TextEncodableEx = True
        Exit Function
    End If
    On Error GoTo bad
    EncodeText s, enc, b
    If ByteCount(b) <= 0 Then Exit Function
    DecodeBytes b, enc, back
    If back = s Then
        TextEncodableEx = True
        Exit Function
    End If
    ' locate the first difference cheaply: 4096-char blocks, then one block char-wise
    blk = 4096
    lo = 1
    Do While lo <= Len(s) And lo <= Len(back)
        hi = lo + blk - 1
        If hi > Len(s) Then hi = Len(s)
        If hi > Len(back) Then hi = Len(back)
        If StrComp(Mid$(s, lo, hi - lo + 1), Mid$(back, lo, hi - lo + 1), vbBinaryCompare) <> 0 Then Exit Do
        lo = hi + 1
    Loop
    hi = lo + blk - 1
    If hi > Len(s) Then hi = Len(s)
    If hi > Len(back) Then hi = Len(back)
    For i = lo To hi
        If Mid$(s, i, 1) <> Mid$(back, i, 1) Then
            firstDiffAt = i
            Exit Function
        End If
    Next i
    Exit Function
bad:
    back = "": firstDiffAt = 0
End Function

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
    ' R-1021-01: a READER must never create anything. MEASURED on the VB6
    ' runtime with an independent probe EXE: "Open <path> For Binary Access
    ' Read" on a MISSING file (parent directory present) does not raise at all -
    ' LOF is 0 and a 0-byte file is now on disk. read_source therefore left
    ' exactly that residue behind and then reported "empty or unreadable".
    ' Probe with GetAttr BEFORE the Open. GetAttr is used ON PURPOSE instead of
    ' CreateObject("Scripting.FileSystemObject"): it needs no COM runtime, it is
    ' cheap (Sha256HexFile / verify hit this often), and above all it FAILS
    ' CLOSED - a missing file is a definite Err 53, never "probe unavailable".
    '   Err 53 (file not found)  -> return the empty array, WITHOUT issuing any Open
    '   any other error (76 = a missing PARENT directory, permission, ...) -> restore
    '   default error handling and let the old path run: it reports the failure
    '   exactly as it always did.
    ' The signature and the "missing/empty => empty array, no error" semantics are
    ' unchanged on purpose: only the CREATE side effect is removed.
    Dim attrProbe As VbFileAttribute
    On Error Resume Next
    Err.Clear
    attrProbe = GetAttr(path)
    If Err.Number = 53 Then
        Err.Clear
        On Error GoTo 0
        Exit Sub
    End If
    If Err.Number <> 0 Then Err.Clear
    On Error GoTo 0
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

' THE CHECKED WRITER. WriteAllBytesTo deliberately swallows its Open error (so a
' failed write used to be invisible), which is the master switch of this whole defect family.
' This variant writes and then READS THE BYTES BACK and compares length + content, so the
' caller can tell 'written' from 'not written'. Callers that report a write to the user must
' use this one (or verify themselves) instead of WriteAllBytesTo.
Public Function WriteAllBytesChecked(ByVal path As String, ByRef b() As Byte) As Boolean
    Dim back() As Byte, i As Long, n As Long, m As Long
    WriteAllBytesChecked = False
    WriteAllBytesTo path, b
    ' read back (ReadAllBytesTo also swallows its error, so an empty/absent file is the signal)
    ReadAllBytesTo path, back
    n = ByteCount(b)
    m = ByteCount(back)
    If m = 0 And n > 0 Then Exit Function
    If m <> n Then Exit Function
    For i = 0 To n - 1
        If back(LBound(back) + i) <> b(LBound(b) + i) Then Exit Function
    Next
    WriteAllBytesChecked = True
End Function

' ############################################################################
' # WARNING: THIS PRIMITIVE DELIBERATELY SWALLOWS ITS Open ERROR.    #
' # A failed write is INVISIBLE to the caller: no error is raised and no value  #
' # is returned. If you are REPORTING A WRITE TO A USER, do NOT call this one -  #
' # call WriteAllBytesChecked() (write + read the bytes back + compare) and      #
' # handle its Boolean, or verify the artifact yourself and refuse on failure.   #
' # tests/check_writeallbytes_callsites.py enforces this: a NEW bare call site    #
' # fails the check.                                                            #
' ############################################################################
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
