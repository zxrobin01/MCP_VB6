VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ServerHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Wires everything together: creates the McpServer + hidden ServerForm and
' adds a "VB6 MCP Status..." entry to the IDE Add-Ins menu.

Private m_vbe As Object
Private m_mcp As McpServer
Private m_form As ServerForm
Private m_wsa As WinsockTransport

' An earlier batch (Planet Source Code 69284 / 54681): a pure-Winsock-API transport that
' does not need MSWINSCK.OCX. Opt-in for now so the hardened OCX path stays the
' default:  VB6MCP_TRANSPORT=winsock  (env)  or  HKCU\Software\VB6MCP\Transport
' = "winsock" (registry). See 内部分析记录
Private m_useWsa As Boolean
Private m_shutdownCalls As Long
Private m_menu As Object
Private WithEvents m_menuEvents As VBIDE.CommandBarEvents
' the second Add-Ins entry - the settings dialog (NON-MODAL).
Private m_menuSettings As Object
Private WithEvents m_settingsEvents As VBIDE.CommandBarEvents

' An earlier batch: IDE change journal sources (signatures taken from the VB6EXT typelib dump,
' 内部类型库转储 -> _dispVBProjectsEvents / _dispVBComponentsEvents /
' _dispContainedVBControlsEvents).
Private WithEvents m_prjEvents As VBIDE.VBProjectsEvents
Private WithEvents m_cmpEvents As VBIDE.VBComponentsEvents
Private WithEvents m_ctlEvents As VBIDE.VBControlsEvents

Public Sub Start(ByVal vbe As Object)
    Set m_vbe = vbe
    Set m_mcp = New McpServer
    m_mcp.Init vbe
    Changes.Reset
    HookIdeEvents
    Dim basePort As Long
    basePort = ReadPortSetting()
    ' only a NAMED port (env/registry) is exclusive; otherwise the
    ' transport probes 8766..8799 for a free one instead of sharing.
    Dim portExplicit As Boolean
    portExplicit = PortSettingExplicit()
    m_useWsa = WantWinsockTransport()
    If m_useWsa Then
        Set m_wsa = New WinsockTransport
        m_wsa.StartListening m_mcp, basePort, portExplicit
    Else
        Set m_form = New ServerForm
        Load m_form
        m_form.StartListening m_mcp, basePort, portExplicit
    End If
    AddMenuItem
    ' install the OPTIONAL code-pane mouse wheel. Default OFF: with
    ' VB6MCP_WHEEL (env) / HKCU\Software\VB6MCP\Wheel unset, nothing at all is
    ' subclassed. No keyboard hook and no global hook is ever installed.
    WinOps.WheelInit
    LogLine "wheel state=" & Left$(WinOps.WheelStateJson(), 500)
    LogLine "handles start=" & WinOps.WheelGuiResourcesJson()
    LogLine "VB6MCP add-in started (requested port " & CStr(basePort) & ")"
End Sub

Public Sub Shutdown()
    m_shutdownCalls = m_shutdownCalls + 1
    LogLine "shutdown enter=" & CStr(m_shutdownCalls)
    ' clean uninstall - restore every code-pane window procedure we
    ' subclassed BEFORE the add-in goes away (a stale pointer would crash the IDE
    ' on exit). This is a no-op when the wheel was never enabled.
    On Error Resume Next
    Dim wheelReceipt As String
    wheelReceipt = WinOps.WheelShutdownJson()
    LogLine "wheel shutdown=" & Left$(wheelReceipt, 700)
    On Error GoTo 0
    UnhookIdeEvents
    RemoveMenuItem
    If Not m_form Is Nothing Then
        m_form.StopListening
        Unload m_form
        Set m_form = Nothing
    End If
    If Not m_wsa Is Nothing Then
        m_wsa.StopListening
        Set m_wsa = Nothing
    End If
    Set m_mcp = Nothing
    Set m_vbe = Nothing
    LogLine "handles end=" & WinOps.WheelGuiResourcesJson()
    LogLine "VB6MCP add-in stopped"
End Sub

Private Function WantWinsockTransport() As Boolean
    Dim s As String
    s = LCase$(Trim$(Environ$("VB6MCP_TRANSPORT")))
    If s = "winsock" Then
        WantWinsockTransport = True
        Exit Function
    End If
    ' advapi32 wrapper - WScript.Shell reads were unreliable in this process
    s = LCase$(Trim$(WinOps.McpSettingGet("Transport")))
    If s = "winsock" Then WantWinsockTransport = True
End Function

' (contract v1.3.1 H2): is the port NAMED by the user (env or registry)?
' A named port must never be replaced silently.
Private Function PortSettingExplicit() As Boolean
    Dim env As String
    env = Environ$("VB6MCP_PORT")
    If Len(env) > 0 Then
        If IsNumeric(env) Then
            PortSettingExplicit = True
            Exit Function
        End If
    End If
    ' advapi32 wrapper
    Dim v As String
    v = WinOps.McpSettingGet("Port")
    If Len(v) > 0 Then
        If IsNumeric(v) Then PortSettingExplicit = True
    End If
End Function

Private Function ReadPortSetting() As Long
    ReadPortSetting = 8766
    Dim env As String
    env = Environ$("VB6MCP_PORT")
    If Len(env) > 0 Then
        If IsNumeric(env) Then
            ReadPortSetting = CLng(Val(env))
            Exit Function
        End If
    End If
    ' advapi32 wrapper
    Dim v As String
    v = WinOps.McpSettingGet("Port")
    If Len(v) > 0 Then
        If IsNumeric(v) Then ReadPortSetting = CLng(v)
    End If
End Function

Private Sub AddMenuItem()
    On Error Resume Next
    Dim addinsMenu As Object
    ' (i18n): Id-first (id=30038) with the caption kept as the
    ' fallback, so a VB6 in any language still finds the Add-Ins menu.
    Set addinsMenu = FindBarControlById(m_vbe, "Menu Bar|" & ChrW$(&H83DC) & ChrW$(&H5355) & ChrW$(&H6761), "Add-Ins|" & ChrW$(&H5916) & ChrW$(&H63A5) & ChrW$(&H7A0B) & ChrW$(&H5E8F), 30038, "")
    If addinsMenu Is Nothing Then Exit Sub
    Set m_menu = addinsMenu.Controls("VB6 MCP Status...")
    If m_menu Is Nothing Then
        Set m_menu = addinsMenu.Controls.Add(1)  ' msoControlButton
        m_menu.Caption = "VB6 MCP Status..."
    End If
    m_menu.Visible = True
    Set m_menuEvents = m_vbe.Events.CommandBarEvents(m_menu)
    ' "VB6 MCP Settings..." - same id-first lookup, own click sink. The settings
    ' form lives in THIS add-in (not in the user's project), so the entry works with no
    ' project open too.
    On Error Resume Next
    Set m_menuSettings = addinsMenu.Controls("VB6 MCP Settings...")
    If m_menuSettings Is Nothing Then
        Set m_menuSettings = addinsMenu.Controls.Add(1)  ' msoControlButton
        m_menuSettings.Caption = "VB6 MCP Settings..."
    End If
    m_menuSettings.Visible = True
    Set m_settingsEvents = m_vbe.Events.CommandBarEvents(m_menuSettings)
End Sub

Private Sub RemoveMenuItem()
    On Error Resume Next
    Set m_menuEvents = Nothing
    If Not m_menu Is Nothing Then
        m_menu.Delete
        Set m_menu = Nothing
    End If
    ' the settings entry and its form go away with the add-in; the form is
    ' unloaded first so no orphan window survives the add-in (the IDE must not crash on exit).
    If Not SettingsForm Is Nothing Then
        If SettingsForm.Visible Then Unload SettingsForm
    End If
    Set m_settingsEvents = Nothing
    If Not m_menuSettings Is Nothing Then
        m_menuSettings.Delete
        Set m_menuSettings = Nothing
    End If
End Sub

' ---------------------------------------------------------------------------
' An earlier batch: IDE change journal. Only events that represent a real model change are
' recorded - Activated/Selected fire constantly and would be pure noise. Handlers
' never let an error escape and only read a Name; they run during run-mode
' transitions and project teardown.
' ---------------------------------------------------------------------------

Private Sub HookIdeEvents()
    On Error Resume Next
    Set m_prjEvents = m_vbe.Events.VBProjectsEvents
    Set m_cmpEvents = m_vbe.Events.VBComponentsEvents(Nothing)
    Set m_ctlEvents = m_vbe.Events.VBControlsEvents(Nothing, Nothing)
    LogLine "IDE event sinks: projects=" & YesNo(m_prjEvents) & " components=" & YesNo(m_cmpEvents) & " controls=" & YesNo(m_ctlEvents)
End Sub

Private Sub UnhookIdeEvents()
    On Error Resume Next
    Set m_prjEvents = Nothing
    Set m_cmpEvents = Nothing
    Set m_ctlEvents = Nothing
End Sub

Private Function YesNo(o As Object) As String
    If o Is Nothing Then YesNo = "no" Else YesNo = "yes"
End Function

Private Function EvName(o As Object) As String
    On Error Resume Next
    EvName = ""
    If o Is Nothing Then Exit Function
    EvName = CStr(o.Name)
End Function

' the instance registry carries the PROJECT identity (and the loaded-build sha, which
' is a load-time constant now). A project event is exactly when that identity can change, so
' tell the watchdog instead of re-reading the VBE object model ~67 times per second.
Private Sub MarkIdentityDirty()
    On Error Resume Next
    If m_form Is Nothing Then Exit Sub
    m_form.MarkIdentityDirty
End Sub

Private Sub m_prjEvents_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    Changes.Record "project_added", EvName(VBProject)
    MarkIdentityDirty
    WinOps.WheelRequestAttach "ide_event"
End Sub

Private Sub m_prjEvents_ItemRemoved(ByVal VBProject As VBIDE.VBProject)
    Changes.Record "project_removed", EvName(VBProject)
    MarkIdentityDirty
    WinOps.WheelRequestAttach "ide_event"
End Sub

Private Sub m_prjEvents_ItemRenamed(ByVal VBProject As VBIDE.VBProject, ByVal OldName As String)
    Changes.Record "project_renamed", OldName & " -> " & EvName(VBProject)
    MarkIdentityDirty
    WinOps.WheelRequestAttach "ide_event"
End Sub

Private Sub m_cmpEvents_ItemAdded(ByVal VBComponent As VBIDE.VBComponent)
    Changes.Record "component_added", EvName(VBComponent)
End Sub

Private Sub m_cmpEvents_ItemRemoved(ByVal VBComponent As VBIDE.VBComponent)
    Changes.Record "component_removed", EvName(VBComponent)
End Sub

Private Sub m_cmpEvents_ItemRenamed(ByVal VBComponent As VBIDE.VBComponent, ByVal OldName As String)
    Changes.Record "component_renamed", OldName & " -> " & EvName(VBComponent)
End Sub

Private Sub m_cmpEvents_ItemReloaded(ByVal VBComponent As VBIDE.VBComponent)
    Changes.Record "component_reloaded", EvName(VBComponent)
End Sub

Private Sub m_ctlEvents_ItemAdded(ByVal VBControl As VBIDE.VBControl)
    Changes.Record "control_added", EvName(VBControl)
End Sub

Private Sub m_ctlEvents_ItemRemoved(ByVal VBControl As VBIDE.VBControl)
    Changes.Record "control_removed", EvName(VBControl)
End Sub

' OldIndex is Long, not Variant: the typelib dump shows it as "Empty", but the
' shipping add-in on the PSC CD (CODE FIXER 53297, Dsr_Connect.Dsr:350) declares
' ByVal OldIndex As Long and compiles.
Private Sub m_ctlEvents_ItemRenamed(ByVal VBControl As VBIDE.VBControl, ByVal OldName As String, ByVal OldIndex As Long)
    Changes.Record "control_renamed", OldName & " -> " & EvName(VBControl)
End Sub

Private Sub m_menuEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    Dim msg As String
    msg = "VB6 MCP Server" & vbCrLf & vbCrLf
    If m_form Is Nothing Then
        msg = msg & "Status: stopped" & vbCrLf
    Else
        msg = msg & "Status: running" & vbCrLf
    End If
    msg = msg & "Transport: " & IIf(m_useWsa, "winsock-api", "MSWINSCK.OCX") & vbCrLf
    msg = msg & "Port: " & CStr(m_mcp.Port) & vbCrLf
    msg = msg & "Project: " & m_mcp.ProjectName()
    MsgBox msg, vbInformation, "VB6 MCP"
End Sub

' open the settings form NON-MODAL, so the MCP server keeps answering while the
' form is up and the IDE is never left blocked (nothing here requires the IDE to be idle).
Private Sub m_settingsEvents_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    On Error Resume Next
    SettingsForm.LivePort = m_mcp.Port
    SettingsForm.BuildSha = m_mcp.ProductBuildSha256()
    SettingsForm.ToolCount = m_mcp.ToolCount()
    If SettingsForm.Visible Then
        SettingsForm.SetFocus
    Else
        SettingsForm.Show vbModeless
    End If
End Sub
