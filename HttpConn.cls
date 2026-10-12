VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HttpConn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Per-socket HTTP/1.1 request accumulator.
' The buffer is a byte-per-char string (see Utf8.bas).

Private Const MAX_REQUEST As Long = 33554432  ' 32 MB

Private m_buf As String
Private m_method As String
Private m_path As String
Private m_headers As Object
Private m_contentLength As Long
Private m_headParsed As Boolean
Private m_error As String
Private m_sse As Boolean

Public Property Get Method() As String
    Method = m_method
End Property

Public Property Get Path() As String
    Path = m_path
End Property

Public Property Get Header(ByVal name As String) As String
    If Not m_headers Is Nothing Then
        If m_headers.Exists(LCase$(name)) Then Header = m_headers.Item(LCase$(name))
    End If
End Property

Public Property Get Body() As String
    ' byte-per-char body
    Dim hdrLen As Long
    hdrLen = HeaderEnd()
    If hdrLen = 0 Then Exit Property
    Body = Mid$(m_buf, hdrLen + 4)
End Property

Public Property Get ErrorText() As String
    ErrorText = m_error
End Property

Public Property Get KeepAlive() As Boolean
    ' Default to closing after each response: clients that never send FIN
    ' would otherwise exhaust the 16-socket pool. Only an explicit
    ' "Connection: keep-alive" request header keeps the socket open.
    If Header("connection") = "keep-alive" Then KeepAlive = True
End Property

Public Property Get IsSse() As Boolean
    IsSse = m_sse
End Property

Public Sub MarkSse()
    m_sse = True
End Sub

Public Sub AppendBytes(b() As Byte)
    On Error Resume Next
    Dim n As Long
    n = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Or n <= 0 Then Exit Sub
    On Error GoTo 0
    If Len(m_buf) + (UBound(b) - LBound(b) + 1) > MAX_REQUEST Then
        m_error = "request too large"
        m_headParsed = True
        Exit Sub
    End If
    m_buf = m_buf & BytesToStr(b)
End Sub

Public Property Get Ready() As Boolean
    If m_error <> "" Then
        Ready = True
        Exit Property
    End If
    If Not m_headParsed Then ParseHead
    If Not m_headParsed Then Exit Property
    Dim idx As Long
    idx = HeaderEnd()
    If idx = 0 Then Exit Property
    If Len(m_buf) - (idx + 3) >= m_contentLength Then Ready = True
End Property

Public Sub AdvanceToNext()
    ' keep any pipelined bytes that follow the current request
    Dim idx As Long
    idx = HeaderEnd()
    If idx = 0 Then
        ClearAll
        Exit Sub
    End If
    If Len(m_buf) >= idx + 3 + m_contentLength Then
        m_buf = Mid$(m_buf, idx + 4 + m_contentLength)
    Else
        m_buf = ""
    End If
    m_method = ""
    m_path = ""
    Set m_headers = Nothing
    m_contentLength = 0
    m_headParsed = False
    m_error = ""
End Sub

Public Sub ClearAll()
    m_buf = ""
    m_method = ""
    m_path = ""
    Set m_headers = Nothing
    m_contentLength = 0
    m_headParsed = False
    m_error = ""
End Sub

Private Function HeaderEnd() As Long
    Dim idx As Long
    idx = InStr(1, m_buf, vbCrLf & vbCrLf)
    If idx = 0 Then idx = InStr(1, m_buf, vbLf & vbLf)
    HeaderEnd = idx
End Function

Private Sub ParseHead()
    If m_headParsed Then Exit Sub
    Dim idx As Long
    idx = HeaderEnd()
    If idx = 0 Then Exit Sub
    m_headParsed = True
    Dim head As String
    head = Left$(m_buf, idx - 1)
    Dim lines() As String
    lines = Split(Replace(head, vbCrLf, vbLf), vbLf)
    Set m_headers = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To UBound(lines)
        Dim ln As String
        ln = lines(i)
        If Len(Trim$(ln)) = 0 Then
            ' skip blank lines
        ElseIf i = 0 Then
            ' request line: METHOD SP PATH SP VERSION
            Dim p1 As Long, p2 As Long
            p1 = InStr(ln, " ")
            If p1 > 0 Then
                m_method = UCase$(Left$(ln, p1 - 1))
                p2 = InStr(p1 + 1, ln, " ")
                If p2 > p1 Then
                    m_path = Mid$(ln, p1 + 1, p2 - p1 - 1)
                Else
                    m_path = Mid$(ln, p1 + 1)
                End If
            End If
        Else
            Dim cpos As Long
            cpos = InStr(ln, ":")
            If cpos > 0 Then
                m_headers(LCase$(Trim$(Left$(ln, cpos - 1)))) = Trim$(Mid$(ln, cpos + 1))
            End If
        End If
    Next
    If m_headers.Exists("content-length") Then m_contentLength = CLng(Val(m_headers("content-length")))
End Sub
