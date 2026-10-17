Attribute VB_Name = "JsonLib"
Option Explicit

' Minimal JSON parser / serializer for VB6MCP (no external references).
' JSON objects -> Scripting.Dictionary (created via CreateObject; all object
'                 traffic flows through Set + object-typed variables only!)
' JSON arrays  -> native VB6 Collection (New Collection)
' JSON null    -> Null
' JSON numbers -> Double (via Val)
'
' CRITICAL (verified on this VB6 build): objects must NEVER cross a
' Variant-Let assignment ("v = obj", "d(k) = objVariant") - it crashes the
' process silently. Safe patterns: "Set x = obj", "Set d(k) = obj",
' "Set x = d.Item(k)", "If IsObject(expr) Then Set x = expr", Collection.Add.

Public Function JsonParse(ByVal s As String, ByRef outValue As Variant) As Boolean
    Dim pos As Long
    pos = 1
    If Len(s) > 0 Then
        If Left$(s, 1) = ChrW$(&HFEFF) Then pos = 2
    End If
    Dim v As Variant
    Dim o As Object
    Dim isObj As Boolean
    If Not ParseValue(s, pos, v, o, isObj) Then Exit Function
    If isObj Then
        Set outValue = o
    Else
        outValue = v
    End If
    JsonParse = True
End Function

Public Function JsonStringify(ByVal v As Variant) As String
    JsonStringify = Stringify(v)
End Function

Public Function JsonGet(ByVal d As Object, ByVal key As String, ByRef outValue As Variant, ByRef outIsObj As Boolean) As Boolean
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsObject(d.Item(key)) Then
        Set outValue = d.Item(key)
        outIsObj = True
    Else
        outValue = d.Item(key)
        outIsObj = False
    End If
    JsonGet = True
End Function

Public Function JsonGetStr(ByVal d As Object, ByVal key As String, Optional ByVal def As String) As String
    If d Is Nothing Then
        JsonGetStr = def
        Exit Function
    End If
    If d.Exists(key) Then
        Dim v As Variant
        If Not IsObject(d.Item(key)) Then
            v = d.Item(key)
            If Not IsNull(v) Then
                JsonGetStr = CStr(v)
                Exit Function
            End If
        End If
    End If
    JsonGetStr = def
End Function

Public Function JsonGetLong(ByVal d As Object, ByVal key As String, Optional ByVal def As Long = 0) As Long
    If d Is Nothing Then
        JsonGetLong = def
        Exit Function
    End If
    If d.Exists(key) Then
        Dim v As Variant
        If Not IsObject(d.Item(key)) Then
            v = d.Item(key)
            If Not IsNull(v) Then
                JsonGetLong = CLng(Val(CStr(v)))
                Exit Function
            End If
        End If
    End If
    JsonGetLong = def
End Function

Public Function JsonGetBool(ByVal d As Object, ByVal key As String, Optional ByVal def As Boolean = False) As Boolean
    If d Is Nothing Then
        JsonGetBool = def
        Exit Function
    End If
    If d.Exists(key) Then
        Dim v As Variant
        If Not IsObject(d.Item(key)) Then
            v = d.Item(key)
            If Not IsNull(v) Then
                If VarType(v) = vbBoolean Then
                    JsonGetBool = CBool(v)
                Else
                    Dim s As String
                    s = LCase$(CStr(v))
                    JsonGetBool = (s = "true" Or s = "1")
                End If
            Else
                JsonGetBool = def
            End If
            Exit Function
        End If
    End If
    JsonGetBool = def
End Function

' ---------- parser (objects only via Set + Object-typed refs) ----------

Private Function ParseValue(ByVal s As String, ByRef pos As Long, ByRef outScalar As Variant, ByRef outObj As Object, ByRef outIsObj As Boolean) As Boolean
    SkipWhitespace s, pos
    If pos > Len(s) Then Exit Function
    Dim c As String
    c = Mid$(s, pos, 1)
    If c = "{" Then
        Dim d As Object
        Set d = CreateObject("Scripting.Dictionary")
        If ParseObjectValue(s, pos, d) Then
            Set outObj = d
            outIsObj = True
            ParseValue = True
        End If
    ElseIf c = "[" Then
        Dim col As New Collection
        If ParseArrayValue(s, pos, col) Then
            Set outObj = col
            outIsObj = True
            ParseValue = True
        End If
    ElseIf c = """" Then
        outScalar = ParseStringValue(s, pos)
        outIsObj = False
        ParseValue = True
    ElseIf c = "t" Then
        If Mid$(s, pos, 4) = "true" Then
            pos = pos + 4
            outScalar = True
            outIsObj = False
            ParseValue = True
        End If
    ElseIf c = "f" Then
        If Mid$(s, pos, 5) = "false" Then
            pos = pos + 5
            outScalar = False
            outIsObj = False
            ParseValue = True
        End If
    ElseIf c = "n" Then
        If Mid$(s, pos, 4) = "null" Then
            pos = pos + 4
            outScalar = Null
            outIsObj = False
            ParseValue = True
        End If
    Else
        outScalar = ParseNumberValue(s, pos)
        outIsObj = False
        ParseValue = True
    End If
End Function

Private Function ParseObjectValue(ByVal s As String, ByRef pos As Long, ByRef d As Object) As Boolean
    pos = pos + 1
    SkipWhitespace s, pos
    If pos <= Len(s) Then
        If Mid$(s, pos, 1) = "}" Then
            pos = pos + 1
            ParseObjectValue = True
            Exit Function
        End If
    End If
    Do
        SkipWhitespace s, pos
        If pos > Len(s) Then Exit Function
        Dim key As String
        key = ParseStringValue(s, pos)
        SkipWhitespace s, pos
        If pos > Len(s) Then Exit Function
        If Mid$(s, pos, 1) <> ":" Then Exit Function
        pos = pos + 1
        Dim val As Variant
        Dim valObj As Object
        Dim valIsObj As Boolean
        If Not ParseValue(s, pos, val, valObj, valIsObj) Then Exit Function
        If valIsObj Then
            Set d(key) = valObj
        Else
            d(key) = val
        End If
        SkipWhitespace s, pos
        If pos <= Len(s) Then
            If Mid$(s, pos, 1) = "," Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    SkipWhitespace s, pos
    If pos > Len(s) Then Exit Function
    If Mid$(s, pos, 1) <> "}" Then Exit Function
    pos = pos + 1
    ParseObjectValue = True
End Function

Private Function ParseArrayValue(ByVal s As String, ByRef pos As Long, ByRef c As Collection) As Boolean
    pos = pos + 1
    SkipWhitespace s, pos
    If pos <= Len(s) Then
        If Mid$(s, pos, 1) = "]" Then
            pos = pos + 1
            ParseArrayValue = True
            Exit Function
        End If
    End If
    Do
        Dim val As Variant
        Dim valObj As Object
        Dim valIsObj As Boolean
        If Not ParseValue(s, pos, val, valObj, valIsObj) Then Exit Function
        If valIsObj Then
            c.Add valObj
        Else
            c.Add val
        End If
        SkipWhitespace s, pos
        If pos <= Len(s) Then
            If Mid$(s, pos, 1) = "," Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    SkipWhitespace s, pos
    If pos > Len(s) Then Exit Function
    If Mid$(s, pos, 1) <> "]" Then Exit Function
    pos = pos + 1
    ParseArrayValue = True
End Function

Private Function ParseStringValue(ByVal s As String, ByRef pos As Long) As String
    ' pos points at the opening quote
    ' the old loop accumulated "out = out & ..." once per character,
    ' which is quadratic on a large JSON string (a replace_code request carries the
    ' whole new text). It now writes into ONE preallocated buffer, so the cost is
    ' O(n). The output is byte-identical: the decoded text is never longer than the
    ' escaped input, because every escape sequence decodes to a single character.
    pos = pos + 1
    Dim n As Long
    n = Len(s)
    Dim buf As String, w As Long
    buf = String$(n - pos + 1, vbNullChar)
    w = 0
    Dim cc As Long
    Dim esc As String
    Do While pos <= n
        cc = AscW(Mid$(s, pos, 1))
        If cc = 34 Then
            pos = pos + 1
            ParseStringValue = Left$(buf, w)
            Exit Function
        ElseIf cc = 92 Then
            pos = pos + 1
            If pos > n Then Err.Raise 1001, "JsonLib", "Unterminated escape"
            esc = Mid$(s, pos, 1)
            pos = pos + 1
            Select Case esc
                Case Chr$(34)
                    Mid$(buf, w + 1, 1) = Chr$(34)
                    w = w + 1
                Case Chr$(92)
                    Mid$(buf, w + 1, 1) = Chr$(92)
                    w = w + 1
                Case Chr$(47)
                    Mid$(buf, w + 1, 1) = Chr$(47)
                    w = w + 1
                Case "b"
                    Mid$(buf, w + 1, 1) = Chr$(8)
                    w = w + 1
                Case "f"
                    Mid$(buf, w + 1, 1) = Chr$(12)
                    w = w + 1
                Case "n"
                    Mid$(buf, w + 1, 1) = vbLf
                    w = w + 1
                Case "r"
                    Mid$(buf, w + 1, 1) = vbCr
                    w = w + 1
                Case "t"
                    Mid$(buf, w + 1, 1) = vbTab
                    w = w + 1
                Case "u"
                    If pos + 3 > n Then Err.Raise 1001, "JsonLib", "Bad \\u escape"
                    Mid$(buf, w + 1, 1) = ChrW$(CLng(Val("&H" & Mid$(s, pos, 4))))
                    w = w + 1
                    pos = pos + 4
                Case Else
                    Err.Raise 1001, "JsonLib", "Bad escape character"
            End Select
        Else
            Mid$(buf, w + 1, 1) = Mid$(s, pos, 1)
            w = w + 1
            pos = pos + 1
        End If
    Loop
    Err.Raise 1001, "JsonLib", "Unterminated string"
End Function

Private Function ParseNumberValue(ByVal s As String, ByRef pos As Long) As Variant
    Dim start As Long
    start = pos
    Do While pos <= Len(s)
        Dim c As String
        c = Mid$(s, pos, 1)
        If InStr("0123456789+-.eE", c) = 0 Then Exit Do
        pos = pos + 1
    Loop
    On Error Resume Next
    ParseNumberValue = Val(Mid$(s, start, pos - start))
End Function

Private Sub SkipWhitespace(ByVal s As String, ByRef pos As Long)
    Do While pos <= Len(s)
        Dim cc As Long
        cc = AscW(Mid$(s, pos, 1))
        If cc = 32 Or cc = 9 Or cc = 10 Or cc = 13 Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' ---------- serializer ----------

Private Function Stringify(ByVal v As Variant) As String
    Dim vt As Integer
    vt = VarType(v)
    If (vt And vbArray) = vbArray Then
        Stringify = StringifyArray(v)
    Else
        Select Case vt
            Case vbNull, vbEmpty
                Stringify = "null"
            Case vbBoolean
                Stringify = IIf(CBool(v), "true", "false")
            Case vbString
                Stringify = Quote(CStr(v))
            Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
                Stringify = Num(v)
            Case vbDate
                Stringify = Quote(Format$(v, "yyyy-mm-dd hh:nn:ss"))
            Case vbObject, vbDataObject
                If v Is Nothing Then
                    Stringify = "null"
                Else
                    Dim tmp As Object
                    Set tmp = v
                    If TypeName(tmp) = "Dictionary" Then
                        Stringify = StringifyDict(tmp)
                    ElseIf TypeName(tmp) = "Collection" Then
                        Stringify = StringifyColl(tmp)
                    Else
                        Err.Raise 1002, "JsonLib", "Cannot serialize object of type " & TypeName(tmp)
                    End If
                End If
            Case Else
                Err.Raise 1002, "JsonLib", "Cannot serialize value of VarType " & CStr(vt)
        End Select
    End If
End Function

Private Function Num(ByVal v As Variant) As String
    Dim s As String
    If VarType(v) = vbDouble Or VarType(v) = vbSingle Then
        s = Trim$(Str$(v))
        If Left$(s, 1) = "." Then s = "0" & s
        If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
        Num = s
    Else
        Num = CStr(v)
    End If
End Function

' item 4: Public so WinProbe.Jq can share ONE JSON string quoter. This body
' escapes every control character (< 32) plus " and \\; 0x7F and everything >= 0x20 are
' passed through byte-for-byte, and the empty string still yields "".
Public Function Quote(ByVal s As String) As String
    ' the per-character writer below is quadratic on a big string
    ' (an earlier measurement: ~12.4 s for 1 MB, exponent ~1.95). Keep it exactly as it
    ' is for small strings, and route large ones through the batched O(n) path.
    If Len(s) > 512 Then
        Quote = QuoteBatch(s)
        Exit Function
    End If
    Dim n As Long
    n = Len(s)
    Dim need As Boolean
    need = False
    Dim i As Long
    Dim cc As Long
    For i = 1 To n
        cc = AscW(Mid$(s, i, 1)) And &HFFFF&
        If cc < 32 Or cc = 34 Or cc = 92 Then
            need = True
            Exit For
        End If
    Next
    If Not need Then
        Quote = """" & s & """"
        Exit Function
    End If
    Dim out As String
    out = Space$(n * 6 + 2)
    Mid$(out, 1, 1) = """"
    Dim j As Long
    j = 2
    For i = 1 To n
        cc = AscW(Mid$(s, i, 1)) And &HFFFF&
        Select Case cc
            Case 34
                Mid$(out, j, 2) = "\"""
                j = j + 2
            Case 92
                Mid$(out, j, 2) = "\\"
                j = j + 2
            Case 8
                Mid$(out, j, 2) = "\b"
                j = j + 2
            Case 9
                Mid$(out, j, 2) = "\t"
                j = j + 2
            Case 10
                Mid$(out, j, 2) = "\n"
                j = j + 2
            Case 12
                Mid$(out, j, 2) = "\f"
                j = j + 2
            Case 13
                Mid$(out, j, 2) = "\r"
                j = j + 2
            Case Else
                If cc < 32 Then
                    Mid$(out, j, 6) = "\u" & Right$("0000" & Hex$(cc), 4)
                    j = j + 6
                Else
                    Mid$(out, j, 1) = ChrW$(cc)
                    j = j + 1
                End If
        End Select
    Next
    Mid$(out, j, 1) = """"
    Quote = Left$(out, j)
End Function

' O(n) batched escaping, used by Quote for strings above 512
' characters - the card's option (2), chosen BY MEASUREMENT: the reference
' implementation (the reference module's StringtoString, option (3): one
' piece per character + a single Join) was implemented and measured here at
' 303 ms for a 1.18 MB string, because it allocates one BSTR per character
' (1.18 M elements); the batch Replace passes below measured 84 ms, and the old
' per-character Mid$ buffer write 131 ms. So the Join technique is used where it
' actually wins (see SliceLines in CodeTools.cls: 19,464 ms -> 14 ms) and the
' batch Replace where it wins here. Order matters: the backslash pass MUST run
' first, so the backslashes it inserts and the backslashes of the later escape
' passes are not escaped a second time. All special characters are built with
' Chr$ so this function's own source carries no escape literals.
Private Function QuoteBatch(ByVal s As String) As String
    Dim out As String
    Dim bs As String, qt As String
    bs = Chr$(92)
    qt = Chr$(34)
    If Not NeedsEscape(s) Then
        QuoteBatch = qt & s & qt
        Exit Function
    End If
    out = s
    If InStr(out, bs) > 0 Then out = Replace$(out, bs, bs & bs)
    If InStr(out, qt) > 0 Then out = Replace$(out, qt, bs & qt)
    If InStr(out, Chr$(8)) > 0 Then out = Replace$(out, Chr$(8), bs & "b")
    If InStr(out, Chr$(9)) > 0 Then out = Replace$(out, Chr$(9), bs & "t")
    If InStr(out, Chr$(10)) > 0 Then out = Replace$(out, Chr$(10), bs & "n")
    If InStr(out, Chr$(12)) > 0 Then out = Replace$(out, Chr$(12), bs & "f")
    If InStr(out, Chr$(13)) > 0 Then out = Replace$(out, Chr$(13), bs & "r")
    out = EscapeOtherControl(out, bs)
    QuoteBatch = qt & out & qt
End Function

' the remaining control characters (0-7, 11, 14-31) become the
' u-escape form with 4 hex digits, exactly as the legacy Case Else produced.
' Called after the backslash pass, so the single backslash it inserts is correct.
Private Function EscapeOtherControl(ByVal s As String, ByVal bs As String) As String
    Dim i As Long
    Dim r As String
    Dim c As String
    r = s
    For i = 0 To 31
        If i <> 8 And i <> 9 And i <> 10 And i <> 12 And i <> 13 Then
            c = Chr$(i)
            If InStr(r, c) > 0 Then
                r = Replace$(r, c, bs & "u" & Right$("0000" & Hex$(i), 4))
            End If
        End If
    Next
    EscapeOtherControl = r
End Function

' exactly the legacy pre-scan predicate (cc < 32, cc = 34, cc = 92)
' but evaluated with native InStr scans instead of one AscW(Mid$(...)) per
' character. Returns True as soon as any character needs escaping.
Private Function NeedsEscape(ByVal s As String) As Boolean
    Dim i As Long
    If InStr(s, Chr$(92)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(34)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(8)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(9)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(10)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(12)) > 0 Then NeedsEscape = True: Exit Function
    If InStr(s, Chr$(13)) > 0 Then NeedsEscape = True: Exit Function
    For i = 0 To 7
        If InStr(s, Chr$(i)) > 0 Then NeedsEscape = True: Exit Function
    Next
    If InStr(s, Chr$(11)) > 0 Then NeedsEscape = True: Exit Function
    For i = 14 To 31
        If InStr(s, Chr$(i)) > 0 Then NeedsEscape = True: Exit Function
    Next
End Function

Private Function StringifyDict(ByVal d As Object) As String
    If d.Count = 0 Then
        StringifyDict = "{}"
        Exit Function
    End If
    Dim keys As Variant
    keys = d.Keys
    Dim parts() As String
    ReDim parts(0 To d.Count - 1)
    Dim i As Long
    For i = 0 To d.Count - 1
        parts(i) = Quote(CStr(keys(i))) & ":" & Stringify(d.Item(keys(i)))
    Next
    StringifyDict = "{" & Join(parts, ",") & "}"
End Function

Private Function StringifyColl(ByVal c As Object) As String
    If c.Count = 0 Then
        StringifyColl = "[]"
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To c.Count - 1)
    ' walk with the Collection's enumerator instead of c.Item(i) - VB6 index
    ' access is a linear walk, so the old loop was O(n^2) on big collections (a failure
    ' receipt carrying 27k hits). Byte-identical output: same order, same elements.
    Dim i As Long, v As Variant
    i = 0
    For Each v In c
        parts(i) = Stringify(v)
        i = i + 1
    Next
    StringifyColl = "[" & Join(parts, ",") & "]"
End Function

Private Function StringifyArray(ByVal v As Variant) As String
    Dim lb As Long, ub As Long
    lb = LBound(v)
    ub = UBound(v)
    If ub < lb Then
        StringifyArray = "[]"
        Exit Function
    End If
    Dim parts() As String
    ReDim parts(0 To ub - lb)
    Dim i As Long
    For i = lb To ub
        parts(i - lb) = Stringify(v(i))
    Next
    StringifyArray = "[" & Join(parts, ",") & "]"
End Function
