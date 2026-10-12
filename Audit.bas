Attribute VB_Name = "Audit"
Option Explicit

Private m_cap As Long
Private m_n As Long
Private m_pos As Long
Private m_path As String
Private m_inited As Boolean
Private m_ts() As String
Private m_tool() As String
Private m_out() As String
Private m_det() As String

Public Sub AuditInit(ByVal logPath As String, Optional ByVal capacity As Long = 500)
    m_path = logPath
    m_cap = capacity
    If m_cap < 1 Then m_cap = 1
    ReDim m_ts(0 To m_cap - 1)
    ReDim m_tool(0 To m_cap - 1)
    ReDim m_out(0 To m_cap - 1)
    ReDim m_det(0 To m_cap - 1)
    m_n = 0
    m_pos = 0
    m_inited = True
End Sub

Private Sub EnsureInit()
    If m_inited Then Exit Sub
    AuditInit Environ$("TEMP") & Chr$(92) & "vb6mcp_audit.log", 500
End Sub

Public Sub AuditAdd(ByVal tool As String, ByVal outcome As String, ByVal detail As String)
    EnsureInit
    Dim i As Long, d As String
    i = m_pos
    m_ts(i) = Format$(Now, "yyyy-mm-dd hh:nn:ss")
    m_tool(i) = tool
    m_out(i) = outcome
    d = detail
    d = Replace(d, vbCrLf, " ")
    d = Replace(d, vbCr, " ")
    d = Replace(d, vbLf, " ")
    If Len(d) > 300 Then d = Left$(d, 300)
    m_det(i) = d
    m_pos = (m_pos + 1) Mod m_cap
    If m_n < m_cap Then m_n = m_n + 1
    On Error Resume Next
    Dim fnum As Integer
    fnum = FreeFile
    Open m_path For Append As #fnum
    Print #fnum, m_ts(i) & vbTab & tool & vbTab & m_out(i) & vbTab & m_det(i)
    Close #fnum
End Sub

Public Function AuditCount() As Long
    AuditCount = m_n
End Function

Public Function AuditTailJson(ByVal n As Long) As String
    EnsureInit
    Dim s As String, idx As Long, cnt As Long
    s = "["
    cnt = 0
    If n > m_n Then n = m_n
    Do While cnt < n
        idx = m_pos - 1 - cnt
        Do While idx < 0
            idx = idx + m_cap
        Loop
        If cnt > 0 Then s = s & ","
        s = s & "{" & Chr$(34) & "ts" & Chr$(34) & ":" & Js(m_ts(idx)) & "," & Chr$(34) & "tool" & Chr$(34) & ":" & Js(m_tool(idx)) & "," & Chr$(34) & "outcome" & Chr$(34) & ":" & Js(m_out(idx)) & "," & Chr$(34) & "detail" & Chr$(34) & ":" & Js(m_det(idx)) & "}"
        cnt = cnt + 1
    Loop
    s = s & "]"
    AuditTailJson = s
End Function

Private Function Js(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Js = Chr$(34) & t & Chr$(34)
End Function
