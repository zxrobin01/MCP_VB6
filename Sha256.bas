Attribute VB_Name = "Sha256"
Option Explicit

Private K(0 To 63) As Long
Private H(0 To 7) As Long
' kernel-probe cache (declarations live here, at the top of the module).
Private mKernelProbed As Boolean
Private mCngOk As Boolean
Private mCapiOk As Boolean

' ---------------------------------------------------------------------------
' (user ruling: MUST DO): the SHA-256 kernel is now the SYSTEM API.
' The old pure-VB6 implementation (64-byte block loop + Double-simulated 32-bit
' arithmetic) took tens of seconds on a 973 KB module, and sha256Before/After are
' computed unconditionally on every edit, which is where the user-visible ~60 s
' replace_code came from. CryptoAPI hashes that input in well under 100 ms and ALL
' call sites benefit because the public signatures are unchanged.
' The VB6 path is KEPT as a bit-identical fallback and can be forced for testing
' with the environment variable VB6MCP_SHA_FALLBACK=1.
' Only advapi32 is used: no COM component and no external exe.
' ---------------------------------------------------------------------------
Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const CALG_SHA_256 As Long = &H800C
Private Const HP_HASHVAL As Long = 2

Private Declare Function CryptAcquireContextA Lib "advapi32" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal hHash As Long, ByRef pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptGetHashParam Lib "advapi32" (ByVal hHash As Long, ByVal dwParam As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal hHash As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

' kernel32 is used for the O(n) UTF-8 conversion (and for the raw
' UTF-16 scan that keeps astral characters byte-identical - see Utf8Bytes below).
Private Const CP_UTF8 As Long = 65001
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cbMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' CNG/bcrypt is the kernel that actually runs. The CryptoAPI path
' installed by never worked on this machine: CryptAcquireContextA
' succeeds but CryptCreateHash(CALG_SHA_256) fails with NTE_BAD_ALGID
' (0x80090008), so every hash silently fell through to the pure-VB6 kernel below
' (~4.3 s per MB). Measured on this machine (CryptoProbe): acquire rc=1 ->
' createhash rc=0 lastErr=0x80090008; forcing VB6MCP_SHA_FALLBACK=1 changes
' nothing. bcrypt.dll hashes 1 MB in ~1 ms and its digest is bit-identical to the
' pure-VB6 kernel (A/B: 7ACC26DF... on the same 1 MB fixture).
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (ByRef phAlgorithm As Long, ByVal pszAlgId As Long, ByVal pszImplementation As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCreateHash Lib "bcrypt" (ByVal hAlgorithm As Long, ByRef phHash As Long, ByVal pbHashObject As Long, ByVal cbHashObject As Long, ByVal pbSecret As Long, ByVal cbSecret As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptHashData Lib "bcrypt" (ByVal hHash As Long, ByRef pbInput As Any, ByVal cbInput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptFinishHash Lib "bcrypt" (ByVal hHash As Long, ByRef pbOutput As Any, ByVal cbOutput As Long, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDestroyHash Lib "bcrypt" (ByVal hHash As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As Long, ByVal dwFlags As Long) As Long

Public Function Sha256HexTextUtf8(ByVal s As String) As String
    Dim u() As Byte
    Erase u
    Utf8Bytes s, u
    Sha256HexTextUtf8 = Sha256HexBytes(u)
End Function

Public Function Sha256HexBytes(ByRef b() As Byte) As String
    ' try the system API first; on ANY failure fall through to the
    ' original pure-VB6 path below, which is byte-for-byte identical (same uppercase
    ' hex, same Utf8Bytes input convention for the text entry point).
    Dim apiOk As Boolean
    If Not ShaForceFallback() Then
        Sha256HexBytes = Sha256HexBytesApi(b, apiOk)
        If apiOk Then Exit Function
    End If
    InitK
    InitH
    Dim n As Long
    n = BCount(b)
    Dim blk(0 To 63) As Byte
    Dim i As Long, j As Long, off As Long
    off = 0
    Do While off + 64 <= n
        MaybePump
        For i = 0 To 63
            blk(i) = b(LBound(b) + off + i)
        Next
        Compress blk
        off = off + 64
    Loop
    Dim remBytes As Long
    remBytes = n - off
    Dim padLen As Long
    padLen = 64
    If remBytes >= 56 Then padLen = 128
    Dim tail() As Byte
    ReDim tail(0 To padLen - 1)
    For i = 0 To remBytes - 1
        MaybePump
        tail(i) = b(LBound(b) + off + i)
    Next
    tail(remBytes) = &H80
    Dim bitLen As Double, hiLen As Double, loLen As Double
    bitLen = CDbl(n) * 8#
    hiLen = Int(bitLen / 4294967296#)
    loLen = bitLen - hiLen * 4294967296#
    PutBE4 tail, padLen - 8, hiLen
    PutBE4 tail, padLen - 4, loLen
    For j = 0 To (padLen \ 64) - 1
        For i = 0 To 63
            blk(i) = tail(j * 64 + i)
        Next
        Compress blk
    Next
    Sha256HexBytes = HashHex()
End Function

' forced-fallback switch. VB6MCP_SHA_FALLBACK=1 makes every call use the
' pure-VB6 implementation, which is how the fallback path is proved identical.
' the settings form exposes this switch as HKCU\Software\VB6MCP\ShaFallback,
' so the fallback now has the same precedence chain the form displays:
'   ENV VB6MCP_SHA_FALLBACK  >  HKCU\Software\VB6MCP\ShaFallback  >  default (off = CryptoAPI)
' The environment still wins, which is what makes the forced-fallback proof reproducible.
Private Function ShaForceFallback() As Boolean
    On Error Resume Next
    Dim s As String
    s = LCase$(Trim$(Environ$("VB6MCP_SHA_FALLBACK")))
    If Len(s) = 0 Then
        ' advapi32 wrapper (WScript.Shell was unreliable in this process)
        s = LCase$(Trim$(WinOps.McpSettingGet("ShaFallback")))
    End If
    ShaForceFallback = (s = "1" Or s = "true" Or s = "on" Or s = "yes")
End Function

' report the kernel that will ACTUALLY run, not the one we wish ran.
' returned a hard-coded "cryptoapi" while the CryptoAPI call was in fact
' failing on every invocation and the pure-VB6 kernel was doing the work - the label
' hid a 4.3 s/MB cost. The probe result is cached in mKernelProbed.
Public Function ShaKernelName() As String
    If ShaForceFallback() Then
        ShaKernelName = "vb6-fallback"
        Exit Function
    End If
    ShaProbeKernels
    If mCngOk Then
        ShaKernelName = "bcrypt-cng"
    ElseIf mCapiOk Then
        ShaKernelName = "cryptoapi"
    Else
        ShaKernelName = "vb6-fallback"
    End If
End Function

Private Sub ShaProbeKernels()
    If mKernelProbed Then Exit Sub
    mKernelProbed = True
    mCngOk = False
    mCapiOk = False
    Dim algName As String
    algName = "SHA256"
    Dim hAlg As Long
    On Error Resume Next
    If BCryptOpenAlgorithmProvider(hAlg, StrPtr(algName), 0, 0) = 0 Then
        mCngOk = True
        BCryptCloseAlgorithmProvider hAlg, 0
    End If
    Dim hProv As Long, hHash As Long
    If CryptAcquireContextA(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT) <> 0 Then
        If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) <> 0 Then
            mCapiOk = True
            CryptDestroyHash hHash
        End If
        CryptReleaseContext hProv, 0
    End If
End Sub

' the kernel. Tries CNG/bcrypt first, then CryptoAPI, then gives up and
' lets the caller fall through to the pure-VB6 implementation below. The digest is
' UPPERCASE hex in every case (identical to HashHex), and the CNG and VB6 kernels are
' bit-identical on the same input (A/B: 7ACC26DF... for the same 1 MB fixture), which
' is what criterion (3) of this card demands.
Private Function Sha256HexBytesApi(ByRef b() As Byte, ByRef ok As Boolean) As String
    Dim n As Long, s As String
    ok = False
    If ShaForceFallback() Then Exit Function
    n = BCount(b)
    If ShaCngHash(b, n, s) Then
        Sha256HexBytesApi = s
        ok = True
        Exit Function
    End If
    If ShaCapiHash(b, n, s) Then
        Sha256HexBytesApi = s
        ok = True
    End If
End Function

Private Function ShaCngHash(ByRef b() As Byte, ByVal n As Long, ByRef hexOut As String) As Boolean
    On Error GoTo Fail
    Dim hAlg As Long, hHash As Long, rc As Long
    Dim algName As String
    algName = "SHA256"
    rc = BCryptOpenAlgorithmProvider(hAlg, StrPtr(algName), 0, 0)
    If rc <> 0 Then Exit Function
    rc = BCryptCreateHash(hAlg, hHash, 0, 0, 0, 0, 0)
    If rc <> 0 Then GoTo Close1
    If n > 0 Then
        rc = BCryptHashData(hHash, b(LBound(b)), n, 0)
    Else
        Dim emptyByte As Byte
        rc = BCryptHashData(hHash, emptyByte, 0, 0)
    End If
    If rc <> 0 Then GoTo Close2
    Dim outBytes(0 To 31) As Byte
    rc = BCryptFinishHash(hHash, outBytes(0), 32, 0)
    If rc = 0 Then
        hexOut = Hex32(outBytes)
        ShaCngHash = True
    End If
Close2:
    BCryptDestroyHash hHash
Close1:
    BCryptCloseAlgorithmProvider hAlg, 0
    Exit Function
Fail:
    ShaCngHash = False
End Function

Private Function ShaCapiHash(ByRef b() As Byte, ByVal n As Long, ByRef hexOut As String) As Boolean
    On Error GoTo Fail
    Dim hProv As Long, hHash As Long, rc As Long
    Dim outBytes(0 To 31) As Byte
    Dim outLen As Long
    rc = CryptAcquireContextA(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)
    If rc = 0 Then Exit Function
    rc = CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash)
    If rc = 0 Then GoTo Rel1
    If n > 0 Then
        rc = CryptHashData(hHash, b(LBound(b)), n, 0)
    Else
        Dim emptyByte As Byte
        rc = CryptHashData(hHash, emptyByte, 0, 0)
    End If
    If rc = 0 Then GoTo Rel2
    outLen = 32
    rc = CryptGetHashParam(hHash, HP_HASHVAL, outBytes(0), outLen, 0)
    If rc <> 0 Then
        If outLen = 32 Then
            hexOut = Hex32(outBytes)
            ShaCapiHash = True
        End If
    End If
Rel2:
    CryptDestroyHash hHash
Rel1:
    CryptReleaseContext hProv, 0
    Exit Function
Fail:
    ShaCapiHash = False
End Function

Private Function Hex32(ByRef b() As Byte) As String
    Dim i As Long, s As String
    For i = 0 To 31
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next
    Hex32 = s
End Function

Private Function BCount(ByRef b() As Byte) As Long
    On Error Resume Next
    BCount = 0
    BCount = UBound(b) - LBound(b) + 1
    On Error GoTo 0
End Function

Private Sub InitK()
    K(0) = &H428A2F98: K(1) = &H71374491: K(2) = &HB5C0FBCF: K(3) = &HE9B5DBA5
    K(4) = &H3956C25B: K(5) = &H59F111F1: K(6) = &H923F82A4: K(7) = &HAB1C5ED5
    K(8) = &HD807AA98: K(9) = &H12835B01: K(10) = &H243185BE: K(11) = &H550C7DC3
    K(12) = &H72BE5D74: K(13) = &H80DEB1FE: K(14) = &H9BDC06A7: K(15) = &HC19BF174
    K(16) = &HE49B69C1: K(17) = &HEFBE4786: K(18) = &H0FC19DC6: K(19) = &H240CA1CC
    K(20) = &H2DE92C6F: K(21) = &H4A7484AA: K(22) = &H5CB0A9DC: K(23) = &H76F988DA
    K(24) = &H983E5152: K(25) = &HA831C66D: K(26) = &HB00327C8: K(27) = &HBF597FC7
    K(28) = &HC6E00BF3: K(29) = &HD5A79147: K(30) = &H06CA6351: K(31) = &H14292967
    K(32) = &H27B70A85: K(33) = &H2E1B2138: K(34) = &H4D2C6DFC: K(35) = &H53380D13
    K(36) = &H650A7354: K(37) = &H766A0ABB: K(38) = &H81C2C92E: K(39) = &H92722C85
    K(40) = &HA2BFE8A1: K(41) = &HA81A664B: K(42) = &HC24B8B70: K(43) = &HC76C51A3
    K(44) = &HD192E819: K(45) = &HD6990624: K(46) = &HF40E3585: K(47) = &H106AA070
    K(48) = &H19A4C116: K(49) = &H1E376C08: K(50) = &H2748774C: K(51) = &H34B0BCB5
    K(52) = &H391C0CB3: K(53) = &H4ED8AA4A: K(54) = &H5B9CCA4F: K(55) = &H682E6FF3
    K(56) = &H748F82EE: K(57) = &H78A5636F: K(58) = &H84C87814: K(59) = &H8CC70208
    K(60) = &H90BEFFFA: K(61) = &HA4506CEB: K(62) = &HBEF9A3F7: K(63) = &HC67178F2
End Sub

Private Sub InitH()
    H(0) = &H6A09E667: H(1) = &HBB67AE85: H(2) = &H3C6EF372: H(3) = &HA54FF53A
    H(4) = &H510E527F: H(5) = &H9B05688C: H(6) = &H1F83D9AB: H(7) = &H5BE0CD19
End Sub

Private Sub Compress(ByRef blk() As Byte)
    Dim w(0 To 63) As Long, t As Long
    For t = 0 To 15
        w(t) = ToLong(CDbl(blk(t * 4)) * 16777216# + CDbl(blk(t * 4 + 1)) * 65536# + CDbl(blk(t * 4 + 2)) * 256# + CDbl(blk(t * 4 + 3)))
    Next
    For t = 16 To 63
        w(t) = Add4(Sig1(w(t - 2)), w(t - 7), Sig0(w(t - 15)), w(t - 16))
    Next
    Dim a As Long, b As Long, c As Long, d As Long, e As Long, f As Long, g As Long, hh As Long
    Dim t1 As Long, t2 As Long
    a = H(0): b = H(1): c = H(2): d = H(3): e = H(4): f = H(5): g = H(6): hh = H(7)
    For t = 0 To 63
        t1 = Add5(hh, BSig1(e), ChF(e, f, g), K(t), w(t))
        t2 = Add32(BSig0(a), MajF(a, b, c))
        hh = g
        g = f
        f = e
        e = Add32(d, t1)
        d = c
        c = b
        b = a
        a = Add32(t1, t2)
    Next
    H(0) = Add32(H(0), a): H(1) = Add32(H(1), b): H(2) = Add32(H(2), c): H(3) = Add32(H(3), d)
    H(4) = Add32(H(4), e): H(5) = Add32(H(5), f): H(6) = Add32(H(6), g): H(7) = Add32(H(7), hh)
End Sub

Private Function Ulg(ByVal x As Long) As Double
    If x >= 0 Then Ulg = CDbl(x) Else Ulg = CDbl(x) + 4294967296#
End Function

Private Function ToLong(ByVal d As Double) As Long
    d = d - Int(d / 4294967296#) * 4294967296#
    If d >= 2147483648# Then d = d - 4294967296#
    ToLong = CLng(d)
End Function

Private Function P2(ByVal s As Long) As Double
    P2 = 2# ^ s
End Function

Private Function SHR(ByVal x As Long, ByVal s As Long) As Long
    SHR = ToLong(Int(Ulg(x) / P2(s)))
End Function

Private Function SHL(ByVal x As Long, ByVal s As Long) As Long
    Dim lo As Double, hi As Double, t As Double
    lo = CDbl(x And &HFFFF&)
    hi = Int(Ulg(x) / 65536#)
    If s < 16 Then
        t = hi * P2(16 + s) + lo * P2(s)
    Else
        t = lo * P2(s)
    End If
    SHL = ToLong(t)
End Function

Private Function ROTR(ByVal x As Long, ByVal s As Long) As Long
    ROTR = (SHR(x, s) Or SHL(x, 32 - s))
End Function

Private Function Add32(ByVal a As Long, ByVal b As Long) As Long
    Add32 = ToLong(Ulg(a) + Ulg(b))
End Function

Private Function Add3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Add3 = ToLong(Ulg(a) + Ulg(b) + Ulg(c))
End Function

Private Function Add4(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    Add4 = ToLong(Ulg(a) + Ulg(b) + Ulg(c) + Ulg(d))
End Function

Private Function Add5(ByVal a As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal e As Long) As Long
    Add5 = ToLong(Ulg(a) + Ulg(b) + Ulg(c) + Ulg(d) + Ulg(e))
End Function

Private Function Sig0(ByVal x As Long) As Long
    Sig0 = (ROTR(x, 7) Xor ROTR(x, 18) Xor SHR(x, 3))
End Function

Private Function Sig1(ByVal x As Long) As Long
    Sig1 = (ROTR(x, 17) Xor ROTR(x, 19) Xor SHR(x, 10))
End Function

Private Function BSig0(ByVal x As Long) As Long
    BSig0 = (ROTR(x, 2) Xor ROTR(x, 13) Xor ROTR(x, 22))
End Function

Private Function BSig1(ByVal x As Long) As Long
    BSig1 = (ROTR(x, 6) Xor ROTR(x, 11) Xor ROTR(x, 25))
End Function

Private Function ChF(ByVal e As Long, ByVal f As Long, ByVal g As Long) As Long
    ChF = ((e And f) Xor ((Not e) And g))
End Function

Private Function MajF(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    MajF = ((a And b) Xor (a And c) Xor (b And c))
End Function

Private Sub PutBE4(ByRef arr() As Byte, ByVal pos As Long, ByVal v As Double)
    Dim x As Double
    x = v
    arr(pos) = CByte(Int(x / 16777216#) - Int(Int(x / 16777216#) / 256#) * 256#)
    x = x - Int(x / 16777216#) * 16777216#
    arr(pos + 1) = CByte(Int(x / 65536#) - Int(Int(x / 65536#) / 256#) * 256#)
    x = x - Int(x / 65536#) * 65536#
    arr(pos + 2) = CByte(Int(x / 256#) - Int(Int(x / 256#) / 256#) * 256#)
    x = x - Int(x / 256#) * 256#
    arr(pos + 3) = CByte(Int(x))
End Sub

Private Function HashHex() As String
    Dim s As String, i As Long
    For i = 0 To 7
        s = s & Right$("00000000" & Hex$(H(i)), 8)
    Next
    HashHex = s
End Function

' O(n) UTF-8 conversion. The loop below used to run for EVERY
' string: one AscW(Mid$(s,i,1)) plus a temporary Byte array of 3n per code unit.
' An earlier measurement took ~3.1 s of that for a 1 MB module, on top of a CryptoAPI kernel
' that needs <10 ms. Now the system converter (kernel32 WideCharToMultiByte /
' CP_UTF8) does the common case in native code.
'
' BIT-IDENTICAL RULE (card criterion 3 - non-negotiable): the old loop encoded
' every UTF-16 code unit SEPARATELY, so a surrogate PAIR (an astral character
' such as an emoji) became two 3-byte sequences (CESU-8), whereas
' WideCharToMultiByte combines the pair into one 4-byte UTF-8 sequence. That would
' change those bytes, so a string containing ANY surrogate code unit still goes
' through the original per-unit loop (exact, and rarely taken); every BMP string
' - i.e. every VB6 source module - takes the system path and is byte-for-byte the
' same as before.
Private Sub Utf8Bytes(ByVal s As String, ByRef o() As Byte)
    Dim n As Long, i As Long, cp As Long, cnt As Long
    n = Len(s)
    If n = 0 Then
        Erase o
        Exit Sub
    End If
    If Utf8BytesSys(s, n, o) Then Exit Sub
    Dim tmp() As Byte
    ReDim tmp(0 To n * 3)
    cnt = 0
    For i = 1 To n
        MaybePump
        cp = AscW(Mid$(s, i, 1))
        If cp < 0 Then cp = cp + 65536
        If cp < &H80 Then
            tmp(cnt) = CByte(cp)
            cnt = cnt + 1
        ElseIf cp < &H800 Then
            tmp(cnt) = CByte(&HC0 Or (cp \ 64))
            tmp(cnt + 1) = CByte(&H80 Or (cp And &H3F))
            cnt = cnt + 2
        Else
            tmp(cnt) = CByte(&HE0 Or (cp \ 4096))
            tmp(cnt + 1) = CByte(&H80 Or ((cp \ 64) And &H3F))
            tmp(cnt + 2) = CByte(&H80 Or (cp And &H3F))
            cnt = cnt + 3
        End If
    Next
    If cnt > 0 Then
        ReDim o(0 To cnt - 1)
        For i = 0 To cnt - 1
            MaybePump
            o(i) = tmp(i)
        Next
    End If
End Sub

' system UTF-8 conversion. Returns False - and the caller falls back
' to the legacy per-unit loop - on a string containing any surrogate code unit, on
' empty/degenerate input, or on any doubt. It never hands back a partial buffer:
' ok is set only after the full declared byte count was produced.
Private Function Utf8BytesSys(ByVal s As String, ByVal n As Long, ByRef o() As Byte) As Boolean
    Dim need As Long, got As Long
    On Error GoTo Fail
    If HasSurrogate(s, n) Then Exit Function
    need = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), n, ByVal 0&, 0, 0, 0)
    If need <= 0 Then Exit Function
    ReDim o(0 To need - 1)
    got = WideCharToMultiByte(CP_UTF8, 0, StrPtr(s), n, o(0), need, 0, 0)
    If got <> need Then Exit Function
    Utf8BytesSys = True
    Exit Function
Fail:
    Utf8BytesSys = False
End Function

' does the string contain any UTF-16 surrogate code unit
' (D800-DFFF)? Checked on the raw UTF-16 bytes: little-endian, so the HIGH byte of
' code unit i lives at the odd offset 2i+1, and D8..DF there is a surrogate. A
' Byte-array scan avoids one Mid$ allocation per character.
Private Function HasSurrogate(ByVal s As String, ByVal n As Long) As Boolean
    Dim b() As Byte, i As Long, hi As Long
    ReDim b(0 To n * 2 - 1)
    CopyMemory b(0), ByVal StrPtr(s), n * 2
    For i = 1 To n * 2 - 1 Step 2
        MaybePump
        hi = b(i)
        If hi >= &HD8 And hi <= &HDF Then
            HasSurrogate = True
            Exit Function
        End If
    Next
End Function
