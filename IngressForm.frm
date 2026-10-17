VERSION 5.00
Begin VB.Form IngressForm 
   Caption         =   "IngressForm"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "IngressForm"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtIngress 
      Height          =   480
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1260
      Visible         =   0   'False
      Width           =   1200
   End
End
Attribute VB_Name = "IngressForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' =====================================================================
' (step 3): the DOORBELL event handler.
'
' READ + CLEAR + ENQUEUE ONLY - it must NEVER execute a tool. The bridge is inside
' SendMessageTimeout(WM_SETTEXT, ..., SMTO_ABORTIFHUNG) while this runs, so a tool
' executed here would hold the bridge for the whole tool duration, which is the
' exact defect this ingress exists to remove. The request is served by
' ServerForm.IngressPump on the watchdog tick, AFTER this handler has returned.
'
' m_inHandler is reset on EVERY exit path (including the error path), and the text
' is cleared BEFORE admission, so:
'   * setting .Text = "" re-enters Change, and the guard returns at once;
'   * a re-entrant call can never see (and re-queue) the same payload.
' =====================================================================
Private m_inHandler As Boolean

Private Sub txtIngress_Change()
    Dim t As String, why As String, detail As String
    If m_inHandler Then Exit Sub
    m_inHandler = True
    On Error GoTo Failed
    t = txtIngress.Text
    If Len(t) > 0 Then
        txtIngress.Text = ""
        Ingress.Accept t, why, detail
        If why <> "ok" Then Ingress.ReportBad t, why, detail
    End If
    m_inHandler = False
    On Error GoTo 0
    Exit Sub
Failed:
    m_inHandler = False
    On Error Resume Next
    Ingress.ReportRaise Err.Number, Err.Description
    On Error GoTo 0
End Sub
