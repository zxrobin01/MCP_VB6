Attribute VB_Name = "ApiSig"
Option Explicit

' C17 - frozen public-API contract: export, spec round-trip, diff.
' Text level only. Parameter NAMES are dropped from the signature so that
' renaming a parameter is not reported as a breaking change.

Private m_ln() As String
Private m_n As Long

Private m_kind() As String
Private m_name() As String
Private m_sig() As String
Private m_cnt As Long

Public Function ApiExportJson(ByVal text As String) As String
    Dim i As Long, out As String
    CollectMembers text
    out = "{" & Qs("count") & ":" & CStr(m_cnt) & "," & Qs("members") & ":["
    For i = 0 To m_cnt - 1
        If i > 0 Then out = out & ","
        out = out & "{" & Qs("kind") & ":" & Js(m_kind(i))
        out = out & "," & Qs("name") & ":" & Js(m_name(i))
        out = out & "," & Qs("sig") & ":" & Js(m_sig(i)) & "}"
    Next
    out = out & "]," & Qs("spec") & ":" & Js(ApiSpecText(text)) & "}"
    ApiExportJson = out
End Function

Public Function ApiSpecText(ByVal text As String) As String
    Dim i As Long, out As String
    CollectMembers text
    For i = 0 To m_cnt - 1
        out = out & m_kind(i) & vbTab & m_name(i) & vbTab & m_sig(i) & vbLf
    Next
    ApiSpecText = out
End Function

Public Function ApiDiffJson(ByVal beforeText As String, ByVal afterText As String) As String
    Dim bk() As String, bn() As String, bs() As String, bc As Long
    Dim ak() As String, an() As String, ag() As String, ac As Long
    CollectMembers beforeText
    SaveSnap bk, bn, bs, bc
    CollectMembers afterText
    SaveSnap ak, an, ag, ac
    ApiDiffJson = DiffCore(bk, bn, bs, bc, ak, an, ag, ac)
End Function

Public Function ApiDiffSpecJson(ByVal specText As String, ByVal afterText As String) As String
    Dim bk() As String, bn() As String, bs() As String, bc As Long
    Dim ak() As String, an() As String, ag() As String, ac As Long
    ParseSpec specText, bk, bn, bs, bc
    CollectMembers afterText
    SaveSnap ak, an, ag, ac
    ApiDiffSpecJson = DiffCore(bk, bn, bs, bc, ak, an, ag, ac)
End Function

Private Function DiffCore(ByRef bk() As String, ByRef bn() As String, ByRef bs() As String, ByVal bc As Long, _
                          ByRef ak() As String, ByRef an() As String, ByRef ag() As String, ByVal ac As Long) As String
    Dim i As Long, j As Long, kkey As String, removed As Long, changed As Long, added As Long
    Dim di As Object
    Set di = CreateObject("Scripting.Dictionary")
    For i = 0 To bc - 1
        kkey = bk(i) & ":" & LCase$(bn(i))
        If Not di.Exists(kkey) Then di.Add kkey, i
    Next
    Dim used() As Boolean
    ReDim used(0 To bc)
    For i = 0 To ac - 1
        kkey = ak(i) & ":" & LCase$(an(i))
        If di.Exists(kkey) Then
            j = di(kkey)
            used(j) = True
            If bs(j) <> ag(i) Then changed = changed + 1
        Else
            added = added + 1
        End If
    Next
    For i = 0 To bc - 1
        If Not used(i) Then removed = removed + 1
    Next
    Dim vo As String, ao As String
    vo = ""
    ao = ""
    For i = 0 To bc - 1
        If Not used(i) Then
            If Len(vo) > 0 Then vo = vo & ","
            vo = vo & "{" & Qs("type") & ":" & Qs("removed") & "," & Qs("kind") & ":" & Js(bk(i))
            vo = vo & "," & Qs("name") & ":" & Js(bn(i)) & "," & Qs("before") & ":" & Js(bs(i)) & "}"
        End If
    Next
    For i = 0 To ac - 1
        kkey = ak(i) & ":" & LCase$(an(i))
        If di.Exists(kkey) Then
            j = di(kkey)
            If bs(j) <> ag(i) Then
                If Len(vo) > 0 Then vo = vo & ","
                vo = vo & "{" & Qs("type") & ":" & Qs("changed") & "," & Qs("kind") & ":" & Js(ak(i))
                vo = vo & "," & Qs("name") & ":" & Js(an(i)) & "," & Qs("before") & ":" & Js(bs(j))
                vo = vo & "," & Qs("after") & ":" & Js(ag(i)) & "}"
            End If
        Else
            If Len(ao) > 0 Then ao = ao & ","
            ao = ao & "{" & Qs("kind") & ":" & Js(ak(i)) & "," & Qs("name") & ":" & Js(an(i)) & "}"
        End If
    Next
    Dim h As String
    h = "{" & Qs("violations") & ":[" & vo & "]," & Qs("added") & ":[" & ao & "]," & Qs("counts") & ":{"
    h = h & Qs("before") & ":" & CStr(bc) & "," & Qs("after") & ":" & CStr(ac)
    h = h & "," & Qs("removed") & ":" & CStr(removed)
    h = h & "," & Qs("changed") & ":" & CStr(changed)
    h = h & "," & Qs("added") & ":" & CStr(added) & "}}"
    DiffCore = h
End Function

Private Sub ParseSpec(ByVal specText As String, ByRef kk() As String, ByRef nn() As String, ByRef ss() As String, ByRef c As Long)
    Dim lines() As String, i As Long, parts() As String, line As String
    lines = Split(specText, vbLf)
    ReDim kk(0 To UBound(lines) + 1)
    ReDim nn(0 To UBound(lines) + 1)
    ReDim ss(0 To UBound(lines) + 1)
    c = 0
    For i = 0 To UBound(lines)
        line = lines(i)
        If Right$(line, 1) = vbCr Then line = Left$(line, Len(line) - 1)
        If Len(line) > 0 Then
            parts = Split(line, vbTab)
            If UBound(parts) >= 2 Then
                kk(c) = parts(0)
                nn(c) = parts(1)
                ss(c) = parts(2)
                c = c + 1
            End If
        End If
    Next
End Sub

Private Sub SaveSnap(ByRef kk() As String, ByRef nn() As String, ByRef ss() As String, ByRef c As Long)
    Dim i As Long
    ReDim kk(0 To m_cnt)
    ReDim nn(0 To m_cnt)
    ReDim ss(0 To m_cnt)
    For i = 0 To m_cnt - 1
        kk(i) = m_kind(i)
        nn(i) = m_name(i)
        ss(i) = m_sig(i)
    Next
    c = m_cnt
End Sub

Private Sub CollectMembers(ByVal text As String)
    Dim i As Long, s As String, depth As Long, kind As String, nm As String, t As String
    BuildLines text
    m_cnt = 0
    ReDim m_kind(0 To m_n)
    ReDim m_name(0 To m_n)
    ReDim m_sig(0 To m_n)
    depth = 0
    For i = 0 To m_n - 1
        s = m_ln(i)
        If Len(s) > 0 Then
            If IsProcHdr(s, kind, nm) Then
                If depth = 0 Then
                    If Not IsPrivateLine(s) Then AddMem kind, nm, NormSig(s)
                End If
                depth = depth + 1
                GoTo Cont
            End If
            If IsProcEnd(s) Then
                If depth > 0 Then depth = depth - 1
                GoTo Cont
            End If
            If depth = 0 Then
                If Not IsPrivateLine(s) Then
                    t = StripMods(s)
                    If Left$(t, 6) = "const " Then
                        AddMem "const", Tok(t, 2), NormSig(s)
                    ElseIf Left$(t, 5) = "type " Then
                        AddMem "type", Tok(t, 2), NormSig(s)
                    ElseIf Left$(t, 5) = "enum " Then
                        AddMem "enum", Tok(t, 2), NormSig(s)
                    ElseIf Left$(t, 6) = "event " Then
                        AddMem "event", Tok(t, 2), NormSig(s)
                    ElseIf Left$(t, 8) = "declare " Then
                        AddMem "declare", Tok(t, 3), NormSig(s)
                    End If
                End If
            End If
        End If
Cont:
    Next
End Sub

Private Sub AddMem(ByVal kind As String, ByVal nm As String, ByVal sig As String)
    If Len(nm) = 0 Then Exit Sub
    m_kind(m_cnt) = kind
    m_name(m_cnt) = nm
    m_sig(m_cnt) = sig
    m_cnt = m_cnt + 1
End Sub

Private Sub BuildLines(ByVal text As String)
    Dim raw() As String, i As Long, cur As String, s As String
    raw = Split(text, vbLf)
    ReDim m_ln(0 To UBound(raw) + 1)
    m_n = 0
    cur = ""
    For i = 0 To UBound(raw)
        s = raw(i)
        If Right$(s, 1) = vbCr Then s = Left$(s, Len(s) - 1)
        s = LCase$(Trim$(Sanitize(s)))
        If Left$(s, 4) = "rem " Then s = ""
        If Len(s) > 0 And Right$(s, 1) = "_" Then
            cur = cur & " " & Left$(s, Len(s) - 1)
        Else
            m_ln(m_n) = Collapse(cur & " " & s)
            m_n = m_n + 1
            cur = ""
        End If
    Next
    If Len(cur) > 0 Then
        m_ln(m_n) = Collapse(cur)
        m_n = m_n + 1
    End If
End Sub

Private Function Sanitize(ByVal s As String) As String
    Dim i As Long, ch As String, inS As Boolean, buf As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inS = Not inS
            buf = buf & " "
        ElseIf inS Then
            buf = buf & " "
        ElseIf ch = Chr$(39) Then
            Exit For
        Else
            buf = buf & ch
        End If
    Next
    Sanitize = buf
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

Private Function IsPrivateLine(ByVal s As String) As Boolean
    If Left$(Trim$(s), 8) = "private " Then IsPrivateLine = True
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

Private Function NormSig(ByVal s As String) As String
    Dim t As String, p1 As Long, p2 As Long, head As String, ps As String, tail As String, out As String, hasP As Boolean
    t = Collapse(s)
    p1 = InStr(t, "(")
    p2 = InStrRev(t, ")")
    If p1 > 0 And p2 > p1 Then
        hasP = True
        head = StripMods(Trim$(Left$(t, p1 - 1)))
        ps = Mid$(t, p1 + 1, p2 - p1 - 1)
        tail = Trim$(Mid$(t, p2 + 1))
    Else
        hasP = False
        head = StripMods(t)
        ps = ""
        tail = ""
    End If
    If hasP Then
        out = head & "(" & NormParams(ps) & ")"
    Else
        If Left$(head, 4) = "sub " Or Left$(head, 9) = "function " Or Left$(head, 9) = "property " Then
            out = head & "()"
        Else
            out = head
        End If
    End If
    If Len(tail) > 0 Then out = out & " " & tail
    NormSig = Collapse(Punctuate(out))
End Function

Private Function NormParams(ByVal s As String) As String
    Dim parts() As String, i As Long, out As String
    If Len(Trim$(s)) = 0 Then NormParams = "" : Exit Function
    parts = Split(s, ",")
    For i = 0 To UBound(parts)
        If i = 0 Then out = NormParam(parts(i)) Else out = out & "," & NormParam(parts(i))
    Next
    NormParams = out
End Function

Private Function NormParam(ByVal p As String) As String
    Dim t As String, mods As String, arr As String, ty As String, q As Long
    t = Trim$(Collapse(p))
    Do
        If Left$(t, 6) = "byval " Then
            mods = mods & "byval " : t = Mid$(t, 7)
        ElseIf Left$(t, 6) = "byref " Then
            mods = mods & "byref " : t = Mid$(t, 7)
        ElseIf Left$(t, 9) = "optional " Then
            mods = mods & "optional " : t = Mid$(t, 10)
        ElseIf Left$(t, 11) = "paramarray " Then
            mods = mods & "paramarray " : t = Mid$(t, 12)
        Else
            Exit Do
        End If
    Loop
    If InStr(t, "()") > 0 Then arr = "()"
    q = InStr(t, " as ")
    If q > 0 Then
        ty = Mid$(t, q + 1)
    Else
        ty = "as variant"
    End If
    NormParam = Collapse(Trim$(mods & arr & " " & ty))
End Function

Private Function Punctuate(ByVal s As String) As String
    Dim t As String
    t = Replace(s, " (", "(")
    t = Replace(t, "( ", "(")
    t = Replace(t, " )", ")")
    t = Replace(t, " ,", ",")
    t = Replace(t, ", ", ",")
    t = Replace(t, " = ", "=")
    t = Replace(t, " =", "=")
    t = Replace(t, "= ", "=")
    Punctuate = t
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
