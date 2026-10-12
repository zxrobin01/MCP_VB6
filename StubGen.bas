Attribute VB_Name = "StubGen"
Option Explicit
' C10 - generate a compilable stub module from a component source text.
' Declarations (Const / Type / Enum / Event / Declare) are kept verbatim;
' public procedures keep their exact header and get an empty body.

Private m_orig() As String
Private m_clean() As String
Private m_raw() As String
Private m_n As Long

Public Function StubSource(ByVal text As String, ByVal moduleName As String) As String
    BuildLines text
    Dim out As String, i As Long, depth As Long, s As String, t As String
    Dim kind As String, nm As String, blockUntil As String
    out = ""
    If Len(moduleName) > 0 Then out = "Attribute VB_Name = " & Chr$(34) & moduleName & Chr$(34) & vbCrLf
    depth = 0
    blockUntil = ""
    For i = 0 To m_n - 1
        MaybePump
        s = m_clean(i)
        If Len(blockUntil) > 0 Then
            If Len(s) > 0 Then out = out & m_raw(i) & vbCrLf
            If s = blockUntil Then blockUntil = ""
            GoTo Cont
        End If
        If Len(s) > 0 Then
            If IsProcHdr(s, kind, nm) Then
                If depth = 0 Then
                    If Not IsPriv(s) Then out = out & m_orig(i) & vbCrLf & EndOf(kind) & vbCrLf
                End If
                depth = depth + 1
                GoTo Cont
            End If
            If IsProcEnd(s) Then
                If depth > 0 Then depth = depth - 1
                GoTo Cont
            End If
            If depth = 0 Then
                If Not IsPriv(s) Then
                    t = StripMods(s)
                    If Left$(t, 5) = "type " Then
                        out = out & m_orig(i) & vbCrLf
                        blockUntil = "end type"
                    ElseIf Left$(t, 5) = "enum " Then
                        out = out & m_orig(i) & vbCrLf
                        blockUntil = "end enum"
                    ElseIf Left$(t, 6) = "const " Then
                        out = out & m_orig(i) & vbCrLf
                    ElseIf Left$(t, 6) = "event " Then
                        out = out & m_orig(i) & vbCrLf
                    ElseIf Left$(t, 8) = "declare " Then
                        out = out & m_orig(i) & vbCrLf
                    End If
                End If
            End If
        End If
Cont:
    Next
    StubSource = out
End Function

Public Function StubJson(ByVal name As String, ByVal text As String) As String
    Dim st As String, n As Long
    st = StubSource(text, name & "_Stub")
    n = UBound(Split(st, vbCrLf))
    If Len(Right$(st, 2)) = 0 Then n = n - 1
    StubJson = "{" & Qs("module") & ":" & Js(name & "_Stub") & "," & Qs("lines") & ":" & CStr(n + 1)
    StubJson = StubJson & "," & Qs("source") & ":" & Js(st) & "}"
End Function

Private Function Qs(ByVal s As String) As String
    Qs = """" & s & """"
End Function

Private Function Js(ByVal s As String) As String
    Dim t As String, b As String
    b = Chr$(92)
    t = Replace(s, vbCrLf, b & "n")
    t = Replace(t, vbLf, b & "n")
    t = Replace(t, vbTab, b & "t")
    t = Replace(t, b, b & b)
    t = Replace(t, """", b & """")
    Js = """" & t & """"
End Function

Private Function EndOf(ByVal kind As String) As String
    If Left$(kind, 8) = "property" Then EndOf = "End Property" : Exit Function
    If kind = "sub" Then EndOf = "End Sub" : Exit Function
    EndOf = "End Function"
End Function

Private Sub BuildLines(ByVal text As String)
    Dim raw() As String, i As Long, cur As String, curC As String, s As String, sc As String, keep As String
    raw = Split(text, vbLf)
    ReDim m_orig(0 To UBound(raw) + 1)
    ReDim m_clean(0 To UBound(raw) + 1)
    ReDim m_raw(0 To UBound(raw) + 1)
    m_n = 0
    cur = ""
    curC = ""
    For i = 0 To UBound(raw)
        MaybePump
        s = raw(i)
        If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        keep = RTrim$(StripComment(s))
        s = Trim$(StripComment(s))
        sc = LCase$(BlankStrings(s))
        If Len(s) > 0 And Right$(s, 1) = "_" Then
            cur = cur & " " & Left$(s, Len(s) - 1)
            curC = curC & " " & Left$(sc, Len(sc) - 1)
        Else
            m_orig(m_n) = Collapse(cur & " " & s)
            m_clean(m_n) = Collapse(curC & " " & sc)
            If Len(cur) > 0 Then m_raw(m_n) = m_orig(m_n) Else m_raw(m_n) = keep
            m_n = m_n + 1
            cur = ""
            curC = ""
        End If
    Next
    If Len(cur) > 0 Then
        m_orig(m_n) = Collapse(cur)
        m_clean(m_n) = Collapse(curC)
        m_raw(m_n) = m_orig(m_n)
        m_n = m_n + 1
    End If
End Sub

Private Function StripComment(ByVal s As String) As String
    Dim i As Long, ch As String, inS As Boolean, buf As String
    For i = 1 To Len(s)
        MaybePump
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inS = Not inS
            buf = buf & ch
        ElseIf inS Then
            buf = buf & ch
        ElseIf ch = Chr$(39) Then
            Exit For
        Else
            buf = buf & ch
        End If
    Next
    StripComment = buf
End Function

Private Function BlankStrings(ByVal s As String) As String
    Dim i As Long, ch As String, inS As Boolean, buf As String
    For i = 1 To Len(s)
        MaybePump
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inS = Not inS
            buf = buf & " "
        ElseIf inS Then
            buf = buf & " "
        Else
            buf = buf & ch
        End If
    Next
    BlankStrings = buf
End Function

Private Function Collapse(ByVal s As String) As String
    Dim t As String
    t = Replace(Trim$(s), vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    Collapse = t
End Function

Private Function StripMods(ByVal s As String) As String
    Dim t As String, again As Boolean
    t = Trim$(s)
    Do
        again = False
        If Left$(t, 7) = "public " Then t = Mid$(t, 8) : again = True
        If (Not again) And Left$(t, 8) = "private " Then t = Mid$(t, 9) : again = True
        If (Not again) And Left$(t, 7) = "friend " Then t = Mid$(t, 8) : again = True
        If (Not again) And Left$(t, 7) = "static " Then t = Mid$(t, 8) : again = True
        If Not again Then Exit Do
    Loop
    StripMods = t
End Function

Private Function IsPriv(ByVal s As String) As Boolean
    If Left$(Trim$(s), 8) = "private " Then IsPriv = True
End Function

Private Function IsProcHdr(ByVal s As String, ByRef kind As String, ByRef nm As String) As Boolean
    Dim t As String, tk As String
    t = StripMods(s)
    nm = ""
    kind = ""
    If Left$(t, 4) = "sub " Then
        kind = "sub" : nm = Tok(t, 2) : IsProcHdr = True : Exit Function
    End If
    If Left$(t, 9) = "function " Then
        kind = "function" : nm = Tok(t, 2) : IsProcHdr = True : Exit Function
    End If
    If Left$(t, 9) = "property " Then
        tk = Tok(t, 2)
        kind = "property:" & tk
        nm = Tok(t, 3)
        IsProcHdr = True
        Exit Function
    End If
End Function

Private Function IsProcEnd(ByVal s As String) As Boolean
    If s = "end sub" Or s = "end function" Or s = "end property" Then IsProcEnd = True
End Function

Private Function Tok(ByVal s As String, ByVal n As Long) As String
    Dim parts() As String, r As String, p As Long
    parts = Split(Trim$(s), " ")
    If n - 1 > UBound(parts) Then Tok = "" : Exit Function
    r = parts(n - 1)
    p = InStr(r, "(")
    If p > 0 Then r = Left$(r, p - 1)
    Tok = r
End Function
