Attribute VB_Name = "McpBpCore"
Option Explicit

' =======================================================================
' C-1 of the CodeTools split: the LEDGER core (BpLedger*) and the dialog/error
' parsing core (H3*). Moved from CodeTools.cls.
'
' Owner convention (docs/DESIGN_CODETOOLS_SPLIT.md section 8): CodeTools has
' VB_PredeclaredId=False, so a module cannot call class methods; every procedure takes
' Owner As CodeTools as its FIRST parameter. Class members used here:
'   ReadTextUtf8File, WriteTextUtf8File, ToolCurrentPosition  (Public)
' The single state field read here (m_bpLedgerStatus) is Public and reached as
' Owner.m_bpLedgerStatus - the state itself stays in the class, single copy.
' =======================================================================

Public Function BpLedgerErrText(Owner As CodeTools, ByVal wr As Object) As String
    BpLedgerErrText = ""
    On Error Resume Next
    If wr Is Nothing Then Exit Function
    BpLedgerErrText = "reason=" & CStr(wr("reason")) & " path=" & CStr(wr("path")) & " bytesExpect=" & CStr(wr("bytesExpect")) & " bytesActual=" & CStr(wr("bytesActual"))
    Err.Clear
End Function

Public Function BpLedgerReason(Owner As CodeTools) As String
    If Owner.m_bpLedgerStatus = "write_invalid" Then
        BpLedgerReason = "ledger_write_invalid"
    Else
        BpLedgerReason = "ledger_write_failed"
    End If
End Function

Public Function BpLedgerStatusOf(Owner As CodeTools, ByVal wr As Object) As String
    BpLedgerStatusOf = "failed"
    On Error Resume Next
    If wr Is Nothing Then Exit Function
    If CBool(wr("ok")) Then
        BpLedgerStatusOf = "ok"
        Exit Function
    End If
    Dim reason As String
    reason = CStr(wr("reason"))
    If reason = "temp_parse_failed" Or reason = "final_parse_failed" Then BpLedgerStatusOf = "write_invalid"
    If reason = "temp_not_object" Or reason = "temp_missing_schema" Then BpLedgerStatusOf = "write_invalid"
    If reason = "temp_missing_breakpoints" Or reason = "temp_count_mismatch" Then BpLedgerStatusOf = "write_invalid"
    If reason = "temp_bytes_mismatch" Or reason = "final_bytes_mismatch" Then BpLedgerStatusOf = "write_invalid"
    Err.Clear
End Function

Public Function BpLedgerStatusStub(Owner As CodeTools, ByVal reason As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = reason
    d("path") = ""
    d("bytesExpect") = 0
    d("bytesActual") = 0
    Set BpLedgerStatusStub = d
End Function

Public Function BpLedgerWriteAtomic(Owner As CodeTools, ByVal p As String, ByVal txt As String, ByVal expectCount As Long) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("reason") = ""
    d("path") = p
    d("bytesExpect") = 0
    d("bytesActual") = 0
    Dim tmp As String
    tmp = p & ".tmp"
    d("tempPath") = tmp
    On Error Resume Next
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    ' FIX-22i: the temp file is DELETED first: WriteTextUtf8File opens For Binary and does NOT
    ' truncate, so a leftover longer temp could otherwise keep its tail.
    If fso.FileExists(tmp) Then fso.DeleteFile tmp, True
    Owner.WriteTextUtf8File tmp, txt
    Dim okR As Boolean, back As String
    back = Owner.ReadTextUtf8File(tmp, okR)
    If Not okR Then
        d("reason") = "temp_read_failed"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    d("bytesExpect") = Len(txt)
    d("bytesActual") = Len(back)
    If StrComp(back, txt, vbBinaryCompare) <> 0 Then
        d("reason") = "temp_bytes_mismatch"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    Dim v As Variant
    If Not JsonParse(back, v) Then
        d("reason") = "temp_parse_failed"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    If Not IsObject(v) Then
        d("reason") = "temp_not_object"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    Dim root As Object
    Set root = v
    If Not root.Exists("schema") Then
        d("reason") = "temp_missing_schema"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    If Not root.Exists("breakpoints") Then
        d("reason") = "temp_missing_breakpoints"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    Dim n As Long
    n = 0
    If IsObject(root("breakpoints")) Then n = root("breakpoints").Count
    If n <> expectCount Then
        d("reason") = "temp_count_mismatch"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    If fso.FileExists(p) Then
        If fso.FileExists(p & ".bak") Then fso.DeleteFile p & ".bak", True
        fso.CopyFile p, p & ".bak", True
    End If
    If fso.FileExists(p) Then fso.DeleteFile p, True
    If Err.Number <> 0 Then
        d("reason") = "delete_target_failed"
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    fso.MoveFile tmp, p
    If Err.Number <> 0 Then
        d("reason") = "rename_failed"
        If fso.FileExists(p & ".bak") Then fso.CopyFile p & ".bak", p, True
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    back = Owner.ReadTextUtf8File(p, okR)
    d("bytesActual") = Len(back)
    If Not okR Then
        d("reason") = "final_read_failed"
        If fso.FileExists(p & ".bak") Then fso.CopyFile p & ".bak", p, True
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    If StrComp(back, txt, vbBinaryCompare) <> 0 Then
        d("reason") = "final_bytes_mismatch"
        If fso.FileExists(p & ".bak") Then fso.CopyFile p & ".bak", p, True
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    If Not JsonParse(back, v) Then
        d("reason") = "final_parse_failed"
        If fso.FileExists(p & ".bak") Then fso.CopyFile p & ".bak", p, True
        Err.Clear
        Set BpLedgerWriteAtomic = d
        Exit Function
    End If
    d("ok") = True
    d("reason") = "ok"
    Set BpLedgerWriteAtomic = d
    Err.Clear
End Function

Public Function H3DialogBlock(Owner As CodeTools, ByVal title As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("present") = True
    d("title") = title
    d("body") = WinOps.DialogBodiesText("#32770")
    Set d("buttons") = H3DialogButtons(Owner)
    Set H3DialogBlock = d
End Function

Public Function H3DialogButtons(Owner As CodeTools) As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    Dim rj As String
    rj = WinOps.ReadActiveDialogsJson("#32770", True, 3)
    Dim v As Variant
    If Not JsonParse(rj, v) Then Set H3DialogButtons = out : Exit Function
    Dim top As Object
    Set top = v
    If top Is Nothing Then Set H3DialogButtons = out : Exit Function
    If Not top.Exists("dialogs") Then Set H3DialogButtons = out : Exit Function
    Dim arr As Object
    Set arr = Nothing
    Set arr = top("dialogs")
    If arr Is Nothing Then Set H3DialogButtons = out : Exit Function
    Dim dlg As Variant
    For Each dlg In arr
        If IsObject(dlg) Then
            Dim ctrls As Object
            Set ctrls = Nothing
            If dlg.Exists("controls") Then Set ctrls = dlg("controls")
            If Not ctrls Is Nothing Then
                Dim c As Variant
                For Each c In ctrls
                    If IsObject(c) Then
                        Dim b As Object
                        Set b = CreateObject("Scripting.Dictionary")
                        b("hwnd") = JsonGetLong(c, "hwnd", 0)
                        b("text") = JsonGetStr(c, "text", "")
                        b("cls") = JsonGetStr(c, "cls", "")
                        out.Add b
                    End If
                Next
            End If
        End If
    Next
    Set H3DialogButtons = out
End Function

Public Function H3ParseErrDescription(Owner As CodeTools, ByVal body As String) As String
    H3ParseErrDescription = ""
    On Error Resume Next
    ' FIX-18: the header itself usually carries the description ("Run-time error 11: Division
    ' by zero" / "ÊµÊ±´íÎó '11': ..."), so take the text AFTER the number on the SAME line
    ' first; only then fall back to a following line. Returning "" here means "no description
    ' was parsed", and the caller keeps errDescription=null (never an empty string).
    ' D1 (OPEN-42 section 4): a VB6 runtime-error header QUOTES the error number, so the
    ' text after the digits starts with the CLOSING QUOTE and only then with the colon:
    '   <zh> '5':      /      Run-time error '5':
    ' The previous code stripped only the colon, so the quote survived and the punctuation
    ' fragment "':" was returned as the description. The quote is now stripped BEFORE the colon
    ' (ASCII apostrophe U+0027, LEFT/RIGHT single quotation marks U+2018/U+2019), and a candidate
    ' with no word character is rejected, so the caller honestly keeps errParsed.description
    ' false instead of faking a parsed description.
    Dim lines() As String
    lines = Split(Replace(Replace(body, vbCrLf, vbLf), vbCr, vbLf), vbLf)
    Dim i As Long
    Dim t As String
    Dim marker As Long
    Dim numEnd As Long
    For i = 0 To UBound(lines)
        t = Trim$(lines(i))
        If Len(t) = 0 Then
            ' skip
        Else
            marker = 0
            If InStr(1, t, ChrW$(&H9519) & ChrW$(&H8BEF), vbTextCompare) > 0 Then marker = 1
            If marker = 0 Then
                If InStr(1, t, "error", vbTextCompare) > 0 Then marker = 1
            End If
            If marker = 1 Then
                numEnd = 0
                Dim j As Long
                Dim ch As String
                Dim seenDigit As Boolean
                seenDigit = False
                For j = 1 To Len(t)
                    ch = Mid$(t, j, 1)
                    If ch >= "0" And ch <= "9" Then
                        seenDigit = True
                        numEnd = j
                    ElseIf seenDigit And ch <> " " And ch <> "'" And ch <> ChrW$(&H2019) Then
                        Exit For
                    End If
                Next
                Dim rest As String
                rest = ""
                If numEnd > 0 Then rest = H3StripErrHeaderTail(Trim$(Mid$(t, numEnd + 1)))
                If Len(rest) > 0 Then
                    If H3ErrDescLooksReal(rest) Then
                        H3ParseErrDescription = rest
                        Exit Function
                    End If
                End If
                Dim k As Long
                For k = i + 1 To UBound(lines)
                    rest = Trim$(lines(k))
                    If Len(rest) > 0 Then
                        If H3ErrDescLooksReal(rest) Then
                            H3ParseErrDescription = rest
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
    Next
End Function

Private Function H3StripErrHeaderTail(ByVal s As String) As String
    ' D1 helper: strip the header punctuation that follows the quoted error number, in the
    ' ORDER the box writes it - closing quote(s) FIRST, then colon(s), then surrounding blanks.
    ' Quotes handled: ASCII apostrophe (U+0027), LEFT (U+2018) and RIGHT (U+2019) single quote.
    ' Colons handled: ASCII ':' and FULLWIDTH COLON U+FF1A.
    H3StripErrHeaderTail = ""
    Dim ch As String
    Do While Len(s) > 0
        ch = Left$(s, 1)
        If ch = "'" Or ch = ChrW$(&H2018) Or ch = ChrW$(&H2019) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    s = Trim$(s)
    Do While Len(s) > 0
        ch = Left$(s, 1)
        If ch = ":" Or ch = ChrW$(&HFF1A) Then
            s = Mid$(s, 2)
        Else
            Exit Do
        End If
    Loop
    H3StripErrHeaderTail = Trim$(s)
End Function

Private Function H3ErrDescLooksReal(ByVal s As String) As Boolean
    ' D1 helper: TRUE only when s carries at least one WORD character (ASCII letter/digit or a
    ' non-ASCII character that is not punctuation). A punctuation-only fragment such as "':" is
    ' NOT a description: returning it would make errParsed.description=true on a failed parse.
    H3ErrDescLooksReal = False
    Dim j As Long
    Dim ch As String
    Dim w As Long
    For j = 1 To Len(s)
        ch = Mid$(s, j, 1)
        w = AscW(ch)
        If w < 0 Then w = w + 65536
        If Not H3IsErrPunctChar(ch, w) Then
            H3ErrDescLooksReal = True
            Exit Function
        End If
    Next
End Function

Private Function H3IsErrPunctChar(ByVal ch As String, ByVal w As Long) As Boolean
    ' D1 helper: TRUE when ONE character is punctuation/whitespace, FALSE when it is a word
    ' character. The ranges below are the Unicode blocks that only carry punctuation/symbols in
    ' the languages VB6 localizes its dialogs to; anything else >= 0x80 counts as a word char.
    H3IsErrPunctChar = True
    If (ch >= "0" And ch <= "9") Or (ch >= "A" And ch <= "Z") Or (ch >= "a" And ch <= "z") Then
        H3IsErrPunctChar = False
        Exit Function
    End If
    If w < &H80 Then Exit Function
    If w = &H2018 Or w = &H2019 Or w = &H201C Or w = &H201D Then Exit Function
    If w = &H2013 Or w = &H2014 Or w = &H2026 Or w = &HAB Or w = &HBB Then Exit Function
    If w >= &H3000 And w <= &H303F Then Exit Function
    If w >= &HFE30 And w <= &HFE4F Then Exit Function
    If w >= &HFF00 And w <= &HFFEF Then Exit Function
    H3IsErrPunctChar = False
End Function
Public Function H3ParseErrNumber(Owner As CodeTools, ByVal body As String) As String
    H3ParseErrNumber = ""
    On Error Resume Next
    Dim i As Long
    i = InStr(1, body, ChrW$(&H9519) & ChrW$(&H8BEF), vbTextCompare)
    If i = 0 Then i = InStr(1, body, "error", vbTextCompare)
    If i = 0 Then Exit Function
    Dim j As Long
    Dim ch As String
    Dim digits As String
    j = i
    digits = ""
    Do While j <= Len(body)
        ch = Mid$(body, j, 1)
        If ch >= "0" And ch <= "9" Then
            digits = digits & ch
        ElseIf Len(digits) > 0 Then
            Exit Do
        End If
        j = j + 1
    Loop
    H3ParseErrNumber = digits
End Function

Public Function H3ParseErrSource(Owner As CodeTools, ByVal body As String) As String
    H3ParseErrSource = ""
    On Error Resume Next
    Dim i As Long
    i = InStr(1, body, "Source:", vbTextCompare)
    If i = 0 Then Exit Function
    Dim rest As String
    rest = Mid$(body, i + 7)
    Dim nl As Long
    nl = InStr(rest, vbLf)
    If nl > 0 Then rest = Left$(rest, nl - 1)
    H3ParseErrSource = Trim$(rest)
End Function

Public Function H3PositionObject(Owner As CodeTools) As Object
    On Error Resume Next
    Set H3PositionObject = Nothing
    Dim rj As String
    rj = Owner.ToolCurrentPosition()
    Dim v As Variant
    If JsonParse(rj, v) Then
        Set H3PositionObject = v
    End If
    Err.Clear
End Function

