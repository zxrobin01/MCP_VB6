VERSION 5.00
Begin VB.Form SettingsForm 
   Caption         =   "VB6 MCP Settings"
   ClientHeight    =   4815
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   LinkTopic       =   "Form1"
   ScaleHeight     =   4815
   ScaleWidth      =   7620
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   6550
      TabIndex        =   10
      Top             =   4300
      Width           =   1050
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "Apply"
      Height          =   360
      Left            =   5400
      TabIndex        =   9
      Top             =   4300
      Width           =   1050
   End
   Begin VB.CheckBox chkSha 
      Caption         =   "Force fallback"
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   1350
      Width           =   2600
   End
   Begin VB.TextBox txtLines 
      Height          =   300
      Left            =   4900
      TabIndex        =   3
      Text            =   "3"
      Top             =   740
      Width           =   700
   End
   Begin VB.CheckBox chkWheel 
      Caption         =   "Enable wheel"
      Height          =   240
      Left            =   240
      TabIndex        =   1
      Top             =   430
      Width           =   2600
   End
   Begin VB.Label lblNote 
      Caption         =   $"SettingsForm.frx":0000
      Height          =   420
      Left            =   180
      TabIndex        =   8
      Top             =   3700
      Width           =   7400
   End
   Begin VB.Label lblStatus 
      Height          =   1000
      Left            =   180
      TabIndex        =   7
      Top             =   2600
      Width           =   7400
   End
   Begin VB.Label lblSource 
      Height          =   800
      Left            =   180
      TabIndex        =   6
      Top             =   1720
      Width           =   7400
   End
   Begin VB.Label lblSha 
      Caption         =   "VB6MCP_SHA_FALLBACK - force the pure-VB6 SHA-256 (default: off)"
      Height          =   220
      Left            =   180
      TabIndex        =   4
      Top             =   1100
      Width           =   7400
   End
   Begin VB.Label lblLines 
      Caption         =   "VB6MCP_WHEEL_LINES - lines per notch (1-50)"
      Height          =   220
      Left            =   180
      TabIndex        =   2
      Top             =   760
      Width           =   4600
   End
   Begin VB.Label lblWheel 
      Caption         =   "VB6MCP_WHEEL - code-pane mouse wheel (default: off)"
      Height          =   220
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   7400
   End
End
Attribute VB_Name = "SettingsForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' TASK-P1-079: VB6 MCP settings form (shown NON-MODAL from the Add-Ins menu so the MCP
' server keeps answering while it is open). Item precedence: ENV > HKCU\Software\VB6MCP >
' built-in default. The form writes HKCU only - it never changes the environment of the VB6
' process (that environment belongs to whoever launched VB6).
Private Const REGKEY As String = "Software\VB6MCP"
Public LivePort As Long
Public BuildSha As String
Public ToolCount As Long

Private Function Truthy(ByVal s As String) As Boolean
    s = LCase$(Trim$(s))
    Truthy = (s = "1" Or s = "true" Or s = "on" Or s = "yes")
End Function

Private Function RegGet(ByVal leaf As String) As String
    ' TASK-P1-079: advapi32 wrapper (WScript.Shell RegRead/RegWrite FAILED inside
    ' this add-in process - measured with a diagnostic probe)
    RegGet = Trim$(WinOps.McpSettingGet(leaf))
End Function

Private Function RegPut(ByVal leaf As String, ByVal v As String) As Boolean
    RegPut = WinOps.McpSettingSet(leaf, v)
End Function

' effective value + where it came from: environment > HKCU > default
Private Sub Describe(ByVal envName As String, ByVal leaf As String, ByVal dflt As String, _
                     ByRef eff As String, ByRef src As String)
    eff = Trim$(Environ$(envName))
    If Len(eff) > 0 Then
        src = "environment"
    Else
        eff = RegGet(leaf)
        If Len(eff) > 0 Then
            src = "HKCU"
        Else
            eff = dflt
            src = "default"
        End If
    End If
End Sub

Private Sub LoadValues()
    Dim wEff As String, wSrc As String, lEff As String, lSrc As String, sEff As String, sSrc As String
    Describe "VB6MCP_WHEEL", "Wheel", "0", wEff, wSrc
    chkWheel.value = IIf(Truthy(wEff), 1, 0)
    Describe "VB6MCP_WHEEL_LINES", "WheelLines", "3", lEff, lSrc
    txtLines.text = lEff
    Describe "VB6MCP_SHA_FALLBACK", "ShaFallback", "0", sEff, sSrc
    chkSha.value = IIf(Truthy(sEff), 1, 0)
    lblSource.caption = "WHEEL=" & wEff & " (" & wSrc & ")" & vbCrLf & _
                        "WHEEL_LINES=" & lEff & " (" & lSrc & ")" & vbCrLf & _
                        "SHA_FALLBACK=" & sEff & " (" & sSrc & ")" & vbCrLf & _
                        "items shown as (environment) cannot be changed from this form."
    lblStatus.caption = "build sha256: " & BuildSha & vbCrLf & _
                        "file version " & CStr(App.major) & "." & CStr(App.minor) & "." & CStr(App.Revision) & _
                        " (DLL version resource)" & vbCrLf & _
                        "product version " & VB6MCP_PRODUCT_VERSION & _
                        "   contract " & VB6MCP_CONTRACT_VERSION & "   tools " & CStr(ToolCount) & vbCrLf & _
                        "live port: " & CStr(LivePort) & "   registry: HKCU\Software\VB6MCP"
End Sub

Private Sub Form_Load()
    Me.caption = "VB6 MCP Settings"
    LoadValues
End Sub

Private Sub cmdApply_Click()
    Dim okAll As Boolean
    Dim wEff As String, wSrc As String, sEff As String, sSrc As String
    Dim n As Long
    okAll = True
    If Not RegPut("Wheel", IIf(chkWheel.value = 1, "1", "0")) Then okAll = False
    n = 3
    If IsNumeric(Trim$(txtLines.text)) Then n = CLng(val(txtLines.text))
    If n < 1 Then n = 1
    If n > 50 Then n = 50
    If Not RegPut("WheelLines", CStr(n)) Then okAll = False
    If Not RegPut("ShaFallback", IIf(chkSha.value = 1, "1", "0")) Then okAll = False
    ' runtime effect for the wheel (the same functions the add-in itself uses)
    Describe "VB6MCP_WHEEL", "Wheel", "0", wEff, wSrc
    If Truthy(wEff) Then
        WinOps.WheelInit
    Else
        WinOps.WheelShutdownJson
    End If
    Describe "VB6MCP_SHA_FALLBACK", "ShaFallback", "0", sEff, sSrc
    LoadValues
    lblNote.caption = "Saved to HKCU\Software\VB6MCP. wheel now " & _
        IIf(Truthy(wEff), "ON", "OFF") & " (" & wSrc & "); SHA fallback " & _
        IIf(Truthy(sEff), "ON", "OFF") & " (" & sSrc & "). " & _
        IIf(okAll, "registry writes: OK", "registry write FAILED")
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
