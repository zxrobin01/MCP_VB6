Attribute VB_Name = "Changes"
Option Explicit

' ---------------------------------------------------------------------------
' Project change journal (P0-1, from the Planet Source Code report's P0-1 item).
' Fed by the three VBIDE event sinks hooked in ServerHost.cls:
'     VBProjectsEvents / VBComponentsEvents / VBControlsEvents
' Rationale: the IDE model has no change feed, and polling get_project_info is
' both costly and unreliable in run mode (内部交接说明 10.1 - CommandBars are
' suspended while the program runs). Events are the durable alternative.
'
' Purely string based on purpose: a sink must NEVER query the IDE model from
' inside an event handler (it fires during teardown and during run-mode state
' transitions). ServerHost does the model access and passes plain strings here.
' Bounded ring buffer - the newest MAX_EVENTS entries are retained.
' ---------------------------------------------------------------------------

Private Const MAX_EVENTS As Long = 500

Private m_rev As Long
Private m_kind() As String
Private m_detail() As String
Private m_time() As String
Private m_revOf() As Long
Private m_head As Long
Private m_count As Long
Private m_dropped As Long

Public Sub Reset()
    m_rev = 0
    m_head = 0
    m_count = 0
    m_dropped = 0
    ReDim m_kind(0 To MAX_EVENTS - 1)
    ReDim m_detail(0 To MAX_EVENTS - 1)
    ReDim m_time(0 To MAX_EVENTS - 1)
    ReDim m_revOf(0 To MAX_EVENTS - 1)
End Sub

Public Property Get Revision() As Long
    Revision = m_rev
End Property

Public Sub Record(ByVal kind As String, ByVal detail As String)
    On Error Resume Next
    m_rev = m_rev + 1
    m_kind(m_head) = kind
    m_detail(m_head) = detail
    m_time(m_head) = Format$(Now, "hh:nn:ss")
    m_revOf(m_head) = m_rev
    m_head = m_head + 1
    If m_head >= MAX_EVENTS Then m_head = 0
    If m_count < MAX_EVENTS Then
        m_count = m_count + 1
    Else
        m_dropped = m_dropped + 1
    End If
End Sub

' Returns the retained events with revision > since, oldest first.
Public Function RevisionJson(ByVal since As Long, ByVal clear As Boolean) As String
    Dim s As String
    Dim i As Long, idx As Long, startIdx As Long, emitted As Long
    If m_count < MAX_EVENTS Then
        startIdx = 0
    Else
        startIdx = m_head
    End If
    s = "{"
    s = s & Jq("revision") & ":" & CStr(m_rev)
    s = s & "," & Jq("retained") & ":" & CStr(m_count)
    s = s & "," & Jq("dropped") & ":" & CStr(m_dropped)
    s = s & "," & Jq("events") & ":["
    For i = 0 To m_count - 1
        idx = startIdx + i
        If idx >= MAX_EVENTS Then idx = idx - MAX_EVENTS
        If m_revOf(idx) > since Then
            If emitted > 0 Then s = s & ","
            s = s & "{" & Jq("rev") & ":" & CStr(m_revOf(idx)) _
                & "," & Jq("kind") & ":" & Jq(m_kind(idx)) _
                & "," & Jq("detail") & ":" & Jq(m_detail(idx)) _
                & "," & Jq("at") & ":" & Jq(m_time(idx)) & "}"
            emitted = emitted + 1
        End If
    Next
    s = s & "]"
    s = s & "," & Jq("returned") & ":" & CStr(emitted)
    RevisionJson = s & "}"
    If clear Then m_dropped = 0
End Function

Private Function Jq(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Jq = Chr$(34) & t & Chr$(34)
End Function
