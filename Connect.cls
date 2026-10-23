VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' VB6 add-in entry point. The IDE creates this class (ProgID VB6MCP.Connect)
' and calls IDTExtensibility2_OnConnection at startup.

Implements IDTExtensibility2

Private m_host As ServerHost

Private Sub IDTExtensibility2_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As ext_ConnectMode, ByVal AddInInst As Object, custom() As Variant)
    On Error Resume Next
    Set m_host = New ServerHost
    m_host.Start VBInst
    LogLine "VB6MCP add-in connected (mode " & CStr(ConnectMode) & ")"
End Sub

Private Sub IDTExtensibility2_OnDisconnection(ByVal RemoveMode As ext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    If Not m_host Is Nothing Then
        LogLine "lifecycle hook=ondisconnection shutdown=run"
        m_host.Shutdown
        Set m_host = Nothing
    Else
        LogLine "lifecycle hook=ondisconnection shutdown=skipped reason=already_shutdown"
    End If
    LogLine "VB6MCP add-in disconnected"
End Sub

Private Sub IDTExtensibility2_OnAddInsUpdate(custom() As Variant)
    ' Keep non-empty: VB6 may drop empty interface procedures.
End Sub

Private Sub IDTExtensibility2_OnStartupComplete(custom() As Variant)
    ' Keep non-empty: VB6 may drop empty interface procedures.
End Sub

Private Sub IDTExtensibility2_OnBeginShutdown(custom() As Variant)
    ' This hook fires BEFORE the host tears its windows down, which is the earliest
    ' point where the code-pane subclassing can be removed safely. OnDisconnection
    ' stays as the fallback for the paths that never call this one.
    On Error Resume Next
    If Not m_host Is Nothing Then
        LogLine "lifecycle hook=onbeginshutdown shutdown=run"
        m_host.Shutdown
        Set m_host = Nothing
    Else
        LogLine "lifecycle hook=onbeginshutdown shutdown=skipped reason=already_shutdown"
    End If
End Sub
