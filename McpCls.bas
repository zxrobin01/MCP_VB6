Attribute VB_Name = "McpCls"
Option Explicit

' =======================================================================
' C-3 facade batch: McpCls (moved bodies). The CLASS keeps same-signature Private
' forwarders, so no class call site changes; each facade delegates to this module.
' Owner convention: Owner As CodeTools is the FIRST parameter of every procedure here.
' =======================================================================

Public Function ComponentPaneText(Owner As CodeTools, ByVal comp As VBIDE.VBComponent) As String
    On Error Resume Next
    ComponentPaneText = ""
    If comp Is Nothing Then Exit Function
    If comp.CodeModule Is Nothing Then Exit Function
    Dim n As Long
    n = comp.CodeModule.CountOfLines
    If n <= 0 Then Exit Function
    ComponentPaneText = comp.CodeModule.Lines(1, n)
End Function

Public Function ControlName(Owner As CodeTools, ByVal ctl As Object) As String
    Dim nm As String
    nm = ""
    On Error Resume Next
    nm = ctl.Name
    If Err.Number <> 0 Then
        Err.Clear
        nm = ""
    End If
    On Error GoTo 0
    If Len(nm) = 0 Then
        Dim prp As Object
        On Error Resume Next
        For Each prp In ctl.Properties
            If LCase$(prp.Name) = "name" Then
                nm = CStr(prp.Value)
                Exit For
            End If
        Next
        If Err.Number <> 0 Then nm = ""
        On Error GoTo 0
    End If
    ControlName = nm
End Function

Public Function DgNow(Owner As CodeTools) As Double
    DgNow = 0#
    If Len(Owner.m_diagClock) = 0 Then Exit Function
    DgNow = McpRe.RpClockValue(Owner, McpRe.RpClockRead(Owner, Owner.m_diagClock))
End Function

Public Sub DgSpan(Owner As CodeTools, ByVal procName As String, ByVal t0 As Double)
    If Len(Owner.m_diagClock) = 0 Then Exit Sub
    If t0 <= 0# Then Exit Sub
    McpRe.RpSpanAdd Owner, Owner.m_diagSpans, procName, McpRe.RpClockValue(Owner, McpRe.RpClockRead(Owner, Owner.m_diagClock)) - t0
End Sub

Public Function DictStr(Owner As CodeTools, ByVal d As Object, ByVal k As String) As String
    On Error Resume Next
    DictStr = ""
    If d.Exists(k) Then DictStr = CStr(d(k))
End Function

Public Function LineUbound(Owner As CodeTools, ByRef arr() As String) As Long
    Dim u As Long
    u = -1
    On Error Resume Next
    u = UBound(arr)
    On Error GoTo 0
    LineUbound = u
End Function

Public Function ProcRangeOf(Owner As CodeTools, ByVal cm As VBIDE.CodeModule, ByVal procName As String, ByRef sLine As Long, ByRef eLine As Long) As Boolean
    On Error Resume Next
    ProcRangeOf = False
    Dim kinds(0 To 3) As Long
    kinds(0) = vbext_pk_Proc
    kinds(1) = vbext_pk_Get
    kinds(2) = vbext_pk_Let
    kinds(3) = vbext_pk_Set
    Dim i As Long, st As Long, n As Long
    For i = 0 To 3
        Err.Clear
        st = 0
        st = cm.ProcStartLine(procName, kinds(i))
        If Err.Number = 0 Then
            If st > 0 Then
                Err.Clear
                n = cm.ProcCountLines(procName, kinds(i))
                If Err.Number = 0 Then
                    If n > 0 Then
                        sLine = st
                        eLine = st + n - 1
                        ProcRangeOf = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
End Function

Public Function ProjectFileNameSafe(Owner As CodeTools) As String
    On Error Resume Next
    Dim p As VBIDE.VBProject
    Set p = Owner.GetProject()
    ProjectFileNameSafe = p.FileName
End Function

Public Function SharedHeaderOffset(Owner As CodeTools, ByVal name As String) As Long
    SharedHeaderOffset = 0
    Dim em As Object
    Set em = Owner.PaneDiskMapCached(name)
    If Not em Is Nothing Then
        If CBool(em("resolved")) Then
            SharedHeaderOffset = CLng(em("headerLines"))
            Exit Function
        End If
    End If
    Dim path As String
    path = Owner.DiskPathOf(name, "")
    If Len(path) > 0 Then SharedHeaderOffset = Owner.HeaderOffsetForPath(path, Owner.CompLineCount(name))
End Function

Public Function TextLines(Owner As CodeTools, ByVal text As String) As String()
    TextLines = Split(Replace(Replace(text, vbCrLf, vbLf), vbCr, vbLf), vbLf)
End Function

Public Function Utf8ByteLen(Owner As CodeTools, ByVal s As String) As Long
    Utf8ByteLen = 0
    If Len(s) = 0 Then Exit Function
    Dim b() As Byte
    b = Utf8Encode(s)
    On Error Resume Next
    Utf8ByteLen = UBound(b) - LBound(b) + 1
    If Err.Number <> 0 Then Utf8ByteLen = 0
    On Error GoTo 0
End Function

