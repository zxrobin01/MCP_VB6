Attribute VB_Name = "DepGraph"
Option Explicit
' C18 - text level dependency graph + purity classify.
' Purity is a coarse token based judgement: it answers whether a component
' can be pulled into an offline test host, and it is deliberately conservative.

Private m_nm() As String
Private m_tx() As String
Private m_cnt As Long

' P0-5: procedure level graph state (component, procedure name, body text)
Private m_pc() As String
Private m_pn() As String
Private m_pb() As String
Private m_pcnt As Long

Private m_color() As Long
Private m_stack() As Long
Private m_sp As Long
Private m_seen As Object
Private m_cyc As String

Public Sub DepReset()
    m_cnt = 0
    ReDim m_nm(0 To 63)
    ReDim m_tx(0 To 63)
    m_pcnt = 0
    ReDim m_pc(0 To 63)
    ReDim m_pn(0 To 63)
    ReDim m_pb(0 To 63)
End Sub

Public Sub DepAdd(ByVal name As String, ByVal text As String)
    If m_cnt > UBound(m_nm) Then
        ReDim Preserve m_nm(0 To m_cnt + 64)
        ReDim Preserve m_tx(0 To m_cnt + 64)
    End If
    m_nm(m_cnt) = name
    m_tx(m_cnt) = text
    m_cnt = m_cnt + 1
    ExtractProcs name, text
End Sub

Public Function DepGraphJson() As String
    Dim i As Long, j As Long, out As String
    If m_cnt = 0 Then
        DepGraphJson = "{" & Qs("nodes") & ":[]," & Qs("edges") & ":[]," & Qs("cycles") & ":[]," & Qs("purity") & ":[]," & Qs("procNodes") & ":[]," & Qs("procEdges") & ":[]," & Qs("procCount") & ":0}"
        Exit Function
    End If
    Dim tk() As String
    ReDim tk(0 To m_cnt)
    For i = 0 To m_cnt - 1
        MaybePump
        tk(i) = Toks(Sanitize(m_tx(i)))
    Next
    Dim adj() As Boolean
    ReDim adj(0 To m_cnt - 1, 0 To m_cnt - 1)
    out = "{" & Qs("nodes") & ":["
    For i = 0 To m_cnt - 1
        If i > 0 Then out = out & ","
        out = out & Js(m_nm(i))
    Next
    out = out & "]," & Qs("edges") & ":["
    Dim first As Boolean
    first = True
    For i = 0 To m_cnt - 1
        For j = 0 To m_cnt - 1
            MaybePump
            If i <> j Then
                If HasTok(tk(i), LCase$(m_nm(j))) Then
                    adj(i, j) = True
                    If Not first Then out = out & ","
                    first = False
                    out = out & "{" & Qs("from") & ":" & Js(m_nm(i)) & "," & Qs("to") & ":" & Js(m_nm(j)) & "}"
                End If
            End If
        Next
    Next
    out = out & "]," & Qs("cycles") & ":[" & CyclesJson(adj) & "]," & Qs("purity") & ":["
    For i = 0 To m_cnt - 1
        If i > 0 Then out = out & ","
        out = out & "{" & Qs("name") & ":" & Js(m_nm(i)) & "," & Qs("pure") & ":" & LCase$(CStr(PurityOf(tk(i))))
        out = out & "," & Qs("reasons") & ":[" & ReasonsJson(tk(i)) & "]}"
    Next
    out = out & "],"
    out = out & Qs("procNodes") & ":[" & ProcNodesJson() & "],"
    out = out & Qs("procEdges") & ":[" & ProcEdgesJson() & "],"
    out = out & Qs("procCount") & ":" & CStr(m_pcnt)
    out = out & "}"
    DepGraphJson = out
End Function

' ---------------------------------------------------------------------------
' P0-5: procedure level caller/callee.
' The component level part above only says "component A mentions component B".
' This resolves real call edges: each procedure body is tokenised (strings and
' comments removed by the same Sanitize) and matched against the set of
' procedure names defined anywhere in the project.
' Cost is one dictionary lookup per token, i.e. O(total tokens) - NOT O(n^2) -
' which matters because VB6 is single threaded (see KNOWN_LIMITATIONS).
' The procedure's own declaration line is excluded so same-named procedures in
' different components (Form_Load, Class_Initialize, ...) do not cross-link.
' ---------------------------------------------------------------------------

Private Sub ExtractProcs(ByVal comp As String, ByVal text As String)
    ' TASK-P1-052: the per-procedure text used to be accumulated with
    ' "acc = acc & lines(i) & vbLf" once per line - quadratic in the PROCEDURE size, so
    ' a single huge procedure paid it in full. The boundary is known, so the body is now
    ' ONE Join over the slice plus the trailing LF. Output is bit-identical.
    Dim lines() As String
    Dim i As Long
    Dim low As String, nm As String
    Dim inProc As Boolean
    Dim pStart As Long
    lines = Split(Replace(text, vbCr, ""), vbLf)
    inProc = False
    nm = ""
    pStart = -1
    For i = 0 To UBound(lines)
        MaybePump
        low = LCase$(Trim$(lines(i)))
        If Not inProc Then
            If ProcHeaderName(low, nm) Then
                inProc = True
                pStart = i
            End If
        Else
            If IsProcEnd(low) Then
                AddProc comp, nm, JoinRangeDG(lines, pStart, i) & vbLf
                inProc = False
                nm = ""
                pStart = -1
            End If
        End If
    Next
    ' truncated source can end with a procedure still open: keep what we have
    If inProc Then AddProc comp, nm, JoinRangeDG(lines, pStart, UBound(lines)) & vbLf
End Sub

' TASK-P1-052: one Join over a slice of a line array - the linear replacement for the
' per-line "acc = acc & lines(i) & vbLf" accumulation.
Private Function JoinRangeDG(ByRef a() As String, ByVal lo As Long, ByVal hi As Long) As String
    Dim cnt As Long, i As Long
    JoinRangeDG = ""
    If lo < 0 Then Exit Function
    cnt = hi - lo
    If cnt < 0 Then Exit Function
    Dim k() As String
    ReDim k(0 To cnt)
    For i = 0 To cnt
        MaybePump
        k(i) = a(lo + i)
    Next
    JoinRangeDG = Join(k, vbLf)
End Function

Private Function ProcHeaderName(ByVal low As String, ByRef nm As String) As Boolean
    Dim s As String, p As Long
    s = low
    Do
        If Left$(s, 7) = "public " Then
            s = Mid$(s, 8)
        ElseIf Left$(s, 8) = "private " Then
            s = Mid$(s, 9)
        ElseIf Left$(s, 7) = "friend " Then
            s = Mid$(s, 8)
        ElseIf Left$(s, 7) = "static " Then
            s = Mid$(s, 8)
        Else
            Exit Do
        End If
    Loop
    If Left$(s, 4) = "sub " Then
        s = Mid$(s, 5)
    ElseIf Left$(s, 9) = "function " Then
        s = Mid$(s, 10)
    ElseIf Left$(s, 13) = "property get " Then
        s = Mid$(s, 14)
    ElseIf Left$(s, 13) = "property let " Then
        s = Mid$(s, 14)
    ElseIf Left$(s, 13) = "property set " Then
        s = Mid$(s, 14)
    Else
        Exit Function
    End If
    p = InStr(s, "(")
    If p = 0 Then p = InStr(s, " ")
    If p > 0 Then s = Left$(s, p - 1)
    s = Trim$(s)
    If Len(s) = 0 Then Exit Function
    nm = s
    ProcHeaderName = True
End Function

Private Function IsProcEnd(ByVal low As String) As Boolean
    IsProcEnd = (low = "end sub") Or (low = "end function") Or (low = "end property")
End Function

Private Sub AddProc(ByVal comp As String, ByVal pname As String, ByVal body As String)
    If m_pcnt > UBound(m_pc) Then
        ReDim Preserve m_pc(0 To m_pcnt + 64)
        ReDim Preserve m_pn(0 To m_pcnt + 64)
        ReDim Preserve m_pb(0 To m_pcnt + 64)
    End If
    m_pc(m_pcnt) = comp
    m_pn(m_pcnt) = pname
    m_pb(m_pcnt) = body
    m_pcnt = m_pcnt + 1
End Sub

Private Function ProcNodesJson() As String
    Dim i As Long, out As String
    out = ""
    For i = 0 To m_pcnt - 1
        MaybePump
        If i > 0 Then out = out & ","
        out = out & Js(m_pc(i) & "." & m_pn(i))
    Next
    ProcNodesJson = out
End Function

Private Function ProcEdgesJson() As String
    Dim dict As Object
    Dim i As Long, j As Long, k As Long, oi As Long, nl As Long
    ' NOTE: the token buffer must NOT be called "toks" - VB6 is case insensitive
    ' and a local named toks shadows the module function Toks, so Toks(...) is
    ' then parsed as an index into the String and the compiler reports the very
    ' unhelpful "expected array" (缺少数组).
    Dim out As String, key As String, tkStr As String, body As String
    Dim parts() As String, owners() As String
    Dim lst As String
    Dim first As Boolean
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 0 To m_pcnt - 1
        key = LCase$(m_pn(i))
        If Not dict.Exists(key) Then dict.Add key, ""
        dict(key) = CStr(dict(key)) & CStr(i) & ","
    Next
    out = ""
    first = True
    For i = 0 To m_pcnt - 1
        MaybePump
        body = m_pb(i)
        nl = InStr(body, vbLf)
        If nl > 0 Then body = Mid$(body, nl + 1)
        tkStr = Toks(Sanitize(body))
        parts = Split(tkStr, " ")
        For k = 0 To UBound(parts)
            key = parts(k)
            If Len(key) > 0 Then
                If dict.Exists(key) Then
                    lst = CStr(dict(key))
                    owners = Split(lst, ",")
                    For j = 0 To UBound(owners) - 1
                        If Len(owners(j)) > 0 Then
                            oi = CLng(owners(j))
                            If oi <> i Then
                                If Not first Then out = out & ","
                                first = False
                                out = out & "{" & Qs("from") & ":" & Js(m_pc(i) & "." & m_pn(i)) _
                                    & "," & Qs("to") & ":" & Js(m_pc(oi) & "." & m_pn(oi)) & "}"
                            End If
                        End If
                    Next
                End If
            End If
        Next
    Next
    ProcEdgesJson = out
End Function

Private Function ReasonsJson(ByVal t As String) As String
    Dim out As String
    out = ""
    If HasTok(t, "new") Then AddR out, "new"
    If HasTok(t, "createobject") Then AddR out, "createobject"
    If HasTok(t, "open") Or HasTok(t, "print") Or HasTok(t, "seek") Then AddR out, "fileio"
    If HasTok(t, "declare") Then AddR out, "apideclare"
    If HasTok(t, "shell") Then AddR out, "shell"
    If HasTok(t, "screen") Or HasTok(t, "forms") Or HasTok(t, "load") Or HasTok(t, "unload") Or HasTok(t, "doevents") Then AddR out, "gui"
    ReasonsJson = out
End Function

Private Sub AddR(ByRef out As String, ByVal r As String)
    If Len(out) > 0 Then out = out & ","
    out = out & Js(r)
End Sub

Private Function PurityOf(ByVal t As String) As Boolean
    PurityOf = (Len(ReasonsJson(t)) = 0)
End Function

Private Function CyclesJson(ByRef adj() As Boolean) As String
    Dim i As Long
    m_cyc = ""
    Set m_seen = CreateObject("Scripting.Dictionary")
    ReDim m_color(0 To m_cnt - 1)
    ReDim m_stack(0 To m_cnt)
    m_sp = 0
    For i = 0 To m_cnt - 1
        MaybePump
        If m_color(i) = 0 Then DfsV i, adj
    Next
    CyclesJson = m_cyc
End Function

Private Sub DfsV(ByVal v As Long, ByRef adj() As Boolean)
    Dim w As Long
    m_color(v) = 1
    m_stack(m_sp) = v
    m_sp = m_sp + 1
    For w = 0 To m_cnt - 1
        MaybePump
        If adj(v, w) Then
            If m_color(w) = 1 Then
                AddCycle w
            ElseIf m_color(w) = 0 Then
                DfsV w, adj
            End If
        End If
    Next
    m_sp = m_sp - 1
    m_color(v) = 2
End Sub

Private Sub AddCycle(ByVal startIdx As Long)
    Dim i As Long, p As Long, key As String, seg As String
    p = -1
    For i = 0 To m_sp - 1
        MaybePump
        If m_stack(i) = startIdx Then p = i : Exit For
    Next
    If p < 0 Then Exit Sub
    For i = p To m_sp - 1
        If i > p Then key = key & "|"
        key = key & CStr(m_stack(i))
        If Len(seg) > 0 Then seg = seg & ","
        seg = seg & Js(m_nm(m_stack(i)))
    Next
    key = key & "|" & CStr(startIdx)
    seg = seg & "," & Js(m_nm(startIdx))
    If m_seen.Exists(key) Then Exit Sub
    m_seen.Add key, 1
    If Len(m_cyc) > 0 Then m_cyc = m_cyc & ","
    m_cyc = m_cyc & "[" & seg & "]"
End Sub

Private Function Toks(ByVal s As String) As String
    ' TASK-P1-058: was a per-character "buf = buf & ch" append - O(n^2) on a whole
    ' module (179 s for dependency_graph on a 27,000-line fixture). A pre-sized buffer
    ' plus Mid$ assignment is O(n); the output is byte-identical (same length, token
    ' characters kept, everything else a space).
    Dim i As Long, n As Long, ch As String, buf As String
    n = Len(s)
    If n = 0 Then Exit Function
    buf = String$(n, " ")
    For i = 1 To n
        MaybePump
        ch = Mid$(s, i, 1)
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Or ch = "_" Then
            Mid$(buf, i, 1) = ch
        End If
    Next
    Toks = buf
End Function

Private Function HasTok(ByVal toks As String, ByVal w As String) As Boolean
    HasTok = (" " & toks & " ") Like ("* " & w & " *")
End Function

Private Function Sanitize(ByVal s As String) As String
    ' TASK-P1-058: linear. The old body appended one character at a time to a growing
    ' BSTR (O(n^2) on a whole module). The output is NEVER longer than the input (a
    ' comment run collapses to its single space), so a pre-sized buffer + Mid$
    ' assignment reproduces the exact byte sequence, only in O(n).
    Dim i As Long, n As Long, o As Long, ch As String, inS As Boolean, buf As String
    n = Len(s)
    If n = 0 Then Exit Function
    buf = String$(n, " ")
    i = 1
    Do While i <= n
        MaybePump
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inS = Not inS
            o = o + 1
            i = i + 1
        ElseIf inS Then
            o = o + 1
            i = i + 1
        ElseIf ch = Chr$(39) Then
            Do While i <= n
                If Mid$(s, i, 1) = vbLf Or Mid$(s, i, 1) = vbCr Then Exit Do
                i = i + 1
            Loop
            o = o + 1
            i = i + 1
        Else
            o = o + 1
            Mid$(buf, o, 1) = ch
            i = i + 1
        End If
    Loop
    Sanitize = LCase$(Left$(buf, o))
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
