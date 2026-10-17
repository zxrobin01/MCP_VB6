Attribute VB_Name = "Ingress"
Option Explicit

' =======================================================================
' (step 3): the INBOUND TEXTBOX-INGRESS ("doorbell") codec, payload
' parser and ONE-SLOT QUEUE.
'
' WHY THIS IS A SEPARATE MODULE: the live path needs a hidden window, a named
' pipe, the bridge and a live IDE. The two things that can be tested WITHOUT any
' of that - "does the base64 decoder work" and "does the payload parse" - are
' exactly the two things that were never isolated when step 3 first failed
' ("reason=malformed" could mean either a broken decoder or a handler that never
' ran). They live here so a throwaway VB6 exe can exercise them offline.
'
' WIRE FORMAT (one newline-free ASCII line, base64 of a UTF-8 byte string):
'     vb6mcp1|<seq>|<nonce16>|<raw JSON-RPC request>
'   seq      decimal, monotonically increasing per bridge run; echoed back so a
'            late/duplicate delivery can be identified instead of re-executed
'   nonce16  the 16-char per-IDE pairing credential the add-in publishes in its
'            instance record; anything else is refused BEFORE any execution
'   raw      the untouched JSON-RPC request text
'
' THE DIVISION OF LABOUR THAT KEEPS THE BRIDGE UNBLOCKED:
'   * IngressForm.txtIngress_Change() = READ + CLEAR + ENQUEUE only. It never
'     executes a tool, so the bridge's SendMessageTimeout(WM_SETTEXT) returns in
'     microseconds instead of for the whole tool duration.
'   * ServerForm.IngressPump() (the watchdog tick) = TAKE + SERVE, i.e. the
'     request is executed AFTER the handler returned, over the already
'     authenticated pipe connection; the reply rides the ordinary pipe path.
' =======================================================================

Public Const INGRESS_MAGIC As String = "vb6mcp1|"
Public Const INGRESS_NONCE_LEN As Long = 16
' An inline payload must fit one window message comfortably. Above this the bridge
' is REQUIRED to fall back to the pipe (it measures the base64 length, not the raw
' length, so the two sides agree without a second negotiation).
Public Const INGRESS_MAX_B64 As Long = 12288

Private Declare Function IngressTickCount Lib "kernel32" Alias "GetTickCount" () As Long

' --- the one-slot queue state -----------------------------------------
Private m_pending As String          ' raw JSON-RPC request text, or ""
Private m_pendingSeq As Long
Private m_pendingAt As Long          ' tick when it was queued
Private m_lastSeq As Long            ' replay/dedup watermark
Private m_statOk As Long
Private m_statBad As Long
Private m_statDup As Long
Private m_statNoAuth As Long
Private m_statBusy As Long
Private m_statDropped As Long
Private m_lastWhy As String
Private m_lastSeqSeen As Long
' The host is the LIVE ServerForm instance (the add-in runs on a "New ServerForm", NOT on
' the form's default instance, so a bare ServerForm.X reference inside IngressForm would
' silently resolve to a SECOND, pipeless instance). Set once at startup; the only thing it
' is used for is the ingress log line.
Private m_host As Object
' The per-process pairing credential this doorbell checks against. It is pushed in by the
' host (PipeTransport.NonceValue) instead of being read back through a form reference, so
' there is exactly one identity and no cross-instance surprise.
Private m_expectedNonce As String

' =======================================================================
' Admission: base64 -> UTF-8 -> parse -> nonce -> dedup -> enqueue.
' NEVER executes anything. Every refusal carries FIELD-LEVEL diagnostics so a
' "malformed" verdict can never again be a dead end.
' =======================================================================
Public Sub SetHost(ByVal host As Object)
    Set m_host = host
End Sub

Public Sub SetCredentials(ByVal nonce16 As String)
    m_expectedNonce = nonce16
End Sub

Public Function ExpectedNonce() As String
    ExpectedNonce = m_expectedNonce
End Function

Public Function Accept(ByVal payloadB64 As String, _
                       ByRef why As String, ByRef detail As String) As Boolean
    On Error GoTo Bad
    Accept = False
    why = ""
    detail = FieldDiag(payloadB64, "")
    Dim n As Long
    n = Len(payloadB64)
    If n = 0 Then
        why = "empty_text"
        m_statBad = m_statBad + 1
        Exit Function
    End If
    If n > INGRESS_MAX_B64 Then
        why = "too_large:" & CStr(n)
        m_statBad = m_statBad + 1
        Exit Function
    End If
    Dim bytes() As Byte, decWhy As String
    If Not B64Decode(payloadB64, bytes, decWhy) Then
        why = "b64:" & decWhy
        m_statBad = m_statBad + 1
        Exit Function
    End If
    Dim text As String
    text = Utf8.Utf8DecodeBytes(bytes)
    detail = FieldDiag(payloadB64, text)
    If Len(text) = 0 Then
        why = "utf8_empty"
        m_statBad = m_statBad + 1
        Exit Function
    End If
    Dim seq As Long, nonce As String, reqText As String, pWhy As String
    If Not PayloadParse(text, seq, nonce, reqText, pWhy) Then
        why = "parse:" & pWhy
        m_statBad = m_statBad + 1
        Exit Function
    End If
    m_lastSeqSeen = seq
    ' Fail CLOSED: without an expected credential nothing is admitted. The doorbell is
    ' authenticated with the SAME 16-char per-process pairing credential the pipe
    ' handshake uses, so a caller that does not hold it cannot execute anything.
    If Len(m_expectedNonce) <= 0 Then
        why = "no_expected_nonce"
        m_statNoAuth = m_statNoAuth + 1
        Exit Function
    End If
    If StrComp(nonce, m_expectedNonce, vbBinaryCompare) <> 0 Then
        ' Authentication failed: NOTHING is executed and the refusal is counted.
        why = "nonce_mismatch"
        m_statNoAuth = m_statNoAuth + 1
        Exit Function
    End If
    If seq > 0 And seq <= m_lastSeq Then
        why = "seq_replay:" & CStr(seq)
        m_statDup = m_statDup + 1
        Exit Function
    End If
    If Len(m_pending) > 0 Then
        ' One in flight, always: an existing pending request is NEVER overwritten.
        why = "slot_busy:" & CStr(seq)
        m_statBusy = m_statBusy + 1
        Exit Function
    End If
    If seq < 0 Then
        why = "seq_negative:" & CStr(seq)
        m_statBad = m_statBad + 1
        Exit Function
    End If
    m_pending = reqText
    m_pendingSeq = seq
    m_pendingAt = IngressTickCount()
    If seq > m_lastSeq Then m_lastSeq = seq
    m_statOk = m_statOk + 1
    why = "ok"
    Accept = True
    Exit Function
Bad:
    why = "internal:" & CStr(Err.Number) & ":" & Err.Description
    m_statBad = m_statBad + 1
    Accept = False
End Function

' The one log call an ingress refusal makes. Late-bound on purpose: the host supplies the
' method, so this module keeps no reference to any form type, and an unset host (the
' offline codec exe) simply logs nothing instead of failing.
Public Sub ReportBad(ByVal raw As String, ByVal why As String, ByVal detail As String)
    On Error Resume Next
    If m_host Is Nothing Then Exit Sub
    m_host.IngressNoteBad raw, why, detail
End Sub

Public Sub ReportRaise(ByVal n As Long, ByVal d As String)
    On Error Resume Next
    If m_host Is Nothing Then Exit Sub
    m_host.IngressNoteRaise n, d
End Sub

Public Function HasPending() As Boolean
    HasPending = (Len(m_pending) > 0)
End Function

Public Function PendingAgeMs() As Long
    If Len(m_pending) = 0 Then Exit Function
    PendingAgeMs = IngressTickCount() - m_pendingAt
End Function

' Take-and-clear. Returns "" when the slot is empty, so the caller can never
' serve the same request twice by accident.
Public Function TakePending(ByRef seq As Long, ByRef ageMs As Long) As String
    seq = 0
    ageMs = 0
    If Len(m_pending) = 0 Then Exit Function
    TakePending = m_pending
    seq = m_pendingSeq
    ageMs = IngressTickCount() - m_pendingAt
    m_pending = ""
    m_pendingSeq = 0
End Function

' Dropping is explicit and counted, never silent.
Public Function DropPending(ByVal reason As String) As String
    If Len(m_pending) = 0 Then
        DropPending = "empty"
        Exit Function
    End If
    m_statDropped = m_statDropped + 1
    m_lastWhy = reason
    DropPending = "seq=" & CStr(m_pendingSeq) & " bytes=" & CStr(Len(m_pending)) & " ageMs=" & CStr(IngressTickCount() - m_pendingAt)
    m_pending = ""
    m_pendingSeq = 0
End Function

Public Sub NoteBad(ByVal why As String)
    m_lastWhy = why
End Sub

' The seq watermark is CONNECTION scoped: the bridge generates seq monotonically for
' the life of one bridge run, so a fresh authenticated connection means a fresh
' numbering. Without this a restarted bridge would start at seq=1 again and every
' request would be misread as a replay of the previous session's seq=1.
Public Sub SessionReset()
    m_lastSeq = 0
End Sub

Public Function LastSeqSeen() As Long
    LastSeqSeen = m_lastSeqSeen
End Function

Public Function StatsJson() As String
    StatsJson = "{""pending"":" & CStr(IIf(HasPending(), 1, 0)) & _
                ",""pendingSeq"":" & CStr(m_pendingSeq) & _
                ",""lastSeq"":" & CStr(m_lastSeq) & _
                ",""ok"":" & CStr(m_statOk) & _
                ",""bad"":" & CStr(m_statBad) & _
                ",""dup"":" & CStr(m_statDup) & _
                ",""noAuth"":" & CStr(m_statNoAuth) & _
                ",""busy"":" & CStr(m_statBusy) & _
                ",""dropped"":" & CStr(m_statDropped) & _
                ",""lastWhy"":""" & JsonEsc(m_lastWhy) & """}"
End Function

' =======================================================================
' Field-level diagnostics. ALWAYS the same five facts, so the caller can put
' them in one log line and a reader never has to guess what was received.
' =======================================================================
Private Function FieldDiag(ByVal s As String, ByVal t As String) As String
    Dim i As Long, p As Long, pipes As Long, magicAt As Long, seqP As Long, nonceP As Long
    pipes = 0
    For i = 1 To Len(t)
        If Mid$(t, i, 1) = "|" Then pipes = pipes + 1
    Next
    magicAt = 0
    If Len(t) >= Len(INGRESS_MAGIC) Then
        If Left$(t, Len(INGRESS_MAGIC)) = INGRESS_MAGIC Then magicAt = 1
    End If
    seqP = InStr(1, t, "|", vbBinaryCompare)
    nonceP = 0
    If seqP > 0 Then nonceP = InStr(seqP + 1, t, "|", vbBinaryCompare)
    FieldDiag = "len_s=" & CStr(Len(s)) & " len_t=" & CStr(Len(t)) & _
                " left16=[" & Left$(t, 16) & "]" & _
                " pipes=" & CStr(pipes) & " magic_at=" & CStr(magicAt) & _
                " instr_first=" & CStr(seqP) & " instr_second=" & CStr(nonceP) & _
                " instr_lenpipe=" & CStr(InStr(1, t, "|", vbBinaryCompare))
End Function

Private Function JsonEsc(ByVal s As String) As String
    Dim i As Long, c As Long, o As String
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c = 34 Then
            o = o & "\"""
        ElseIf c = 92 Then
            o = o & "\\"
        ElseIf c < 32 Or c > 126 Then
            o = o & "?"
        Else
            o = o & Chr$(c)
        End If
    Next
    JsonEsc = o
End Function

' Standard-alphabet base64 -> raw bytes, STRICT.
' Every refusal names the exact position/character so a "malformed" verdict can
' never again be a dead end: why is one of
'   empty | length_not_multiple_of_4:<n> | no_payload |
'   bad_char_at_<i>:&H<hex> | padding_not_at_end:<i> | internal:<n>:<desc>
Public Function B64Decode(ByVal s As String, ByRef outBytes() As Byte, ByRef why As String) As Boolean
    On Error GoTo Bad
    B64Decode = False
    why = ""
    Dim n As Long
    n = Len(s)
    If n = 0 Then
        why = "empty"
        Exit Function
    End If
    If (n Mod 4) <> 0 Then
        why = "length_not_multiple_of_4:" & CStr(n)
        Exit Function
    End If
    Dim pad As Long
    pad = 0
    If Right$(s, 2) = "==" Then
        pad = 2
    ElseIf Right$(s, 1) = "=" Then
        pad = 1
    End If
    ' ---- character validation (position-exact) -------------------------
    Dim i As Long, a As Long, ch As String
    For i = 1 To n
        ch = Mid$(s, i, 1)
        a = AscW(ch)
        If a = 61 Then
            If i <= (n - pad) Then
                why = "padding_not_at_end:" & CStr(i)
                Exit Function
            End If
        ElseIf B64Val(ch) < 0 Then
            why = "bad_char_at_" & CStr(i) & ":&H" & Hex$(a)
            Exit Function
        End If
    Next
    Dim outLen As Long
    outLen = (n \ 4) * 3 - pad
    If outLen <= 0 Then
        why = "no_payload"
        Exit Function
    End If
    ReDim outBytes(0 To outLen - 1)
    ' ---- decode --------------------------------------------------------
    Dim blk As Long, base As Long, o As Long
    Dim v0 As Long, v1 As Long, v2 As Long, v3 As Long
    For blk = 0 To (n \ 4) - 1
        base = blk * 4
        v0 = B64Val(Mid$(s, base + 1, 1))
        v1 = B64Val(Mid$(s, base + 2, 1))
        v2 = B64Val(Mid$(s, base + 3, 1))
        v3 = B64Val(Mid$(s, base + 4, 1))
        If v2 < 0 Then v2 = 0
        If v3 < 0 Then v3 = 0
        If o < outLen Then
            outBytes(o) = CByte(((v0 * 4) Or (v1 \ 16)) And &HFF)
            o = o + 1
        End If
        If o < outLen Then
            outBytes(o) = CByte(((v1 * 16) Or (v2 \ 4)) And &HFF)
            o = o + 1
        End If
        If o < outLen Then
            outBytes(o) = CByte(((v2 * 64) Or v3) And &HFF)
            o = o + 1
        End If
    Next
    B64Decode = True
    Exit Function
Bad:
    why = "internal:" & CStr(Err.Number) & ":" & Err.Description
    B64Decode = False
End Function

' One base64 character -> its 6-bit value, or -1 when it is not in the alphabet.
Private Function B64Val(ByVal ch As String) As Long
    Dim a As Long
    a = AscW(ch)
    If a >= 65 And a <= 90 Then
        B64Val = a - 65
    ElseIf a >= 97 And a <= 122 Then
        B64Val = a - 97 + 26
    ElseIf a >= 48 And a <= 57 Then
        B64Val = a - 48 + 52
    ElseIf a = 43 Then
        B64Val = 62
    ElseIf a = 47 Then
        B64Val = 63
    Else
        B64Val = -1
    End If
End Function

' vb6mcp1|<seq>|<nonce16>|<raw> -> the three fields. why is field-exact:
'   too_short:<n> | bad_magic:<first16> | no_seq_pipe | empty_seq |
'   seq_not_numeric:<v> | seq_negative | no_nonce_pipe | nonce_len:<n> |
'   empty_request | internal:<n>:<desc>
Public Function PayloadParse(ByVal text As String, ByRef seq As Long, _
                             ByRef nonce As String, ByRef reqText As String, _
                             ByRef why As String) As Boolean
    On Error GoTo Bad
    PayloadParse = False
    why = ""
    seq = 0
    nonce = ""
    reqText = ""
    If Len(text) < 12 Then
        why = "too_short:" & CStr(Len(text))
        Exit Function
    End If
    If Left$(text, Len(INGRESS_MAGIC)) <> INGRESS_MAGIC Then
        why = "bad_magic:" & Left$(text, 16)
        Exit Function
    End If
    Dim p As Long, q As Long
    p = Len(INGRESS_MAGIC) + 1
    q = InStr(p, text, "|")
    If q = 0 Then
        why = "no_seq_pipe"
        Exit Function
    End If
    Dim sSeq As String
    sSeq = Mid$(text, p, q - p)
    If Len(sSeq) = 0 Then
        why = "empty_seq"
        Exit Function
    End If
    If Not IsNumeric(sSeq) Then
        why = "seq_not_numeric:" & sSeq
        Exit Function
    End If
    If CDbl(sSeq) < 0# Then
        why = "seq_negative"
        Exit Function
    End If
    seq = CLng(sSeq)
    p = q + 1
    q = InStr(p, text, "|")
    If q = 0 Then
        why = "no_nonce_pipe"
        Exit Function
    End If
    nonce = Mid$(text, p, q - p)
    If Len(nonce) <> INGRESS_NONCE_LEN Then
        why = "nonce_len:" & CStr(Len(nonce))
        Exit Function
    End If
    reqText = Mid$(text, q + 1)
    If Len(reqText) = 0 Then
        why = "empty_request"
        Exit Function
    End If
    PayloadParse = True
    Exit Function
Bad:
    why = "internal:" & CStr(Err.Number) & ":" & Err.Description
    PayloadParse = False
End Function
