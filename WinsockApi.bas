Attribute VB_Name = "WinsockApi"
Option Explicit

' ---------------------------------------------------------------------------
' Pure Winsock API transport primitives (no MSWINSCK.OCX).
' Approach borrowed (not code) from two Planet Source Code submissions:
'   69284  data\src\HTTP Class\MSocketSupport.bas  (CreateWinsockMessageWindow/WindowProc)
'   54681  data\srcx\CSocketMas176430712004\...     (CSocketMaster; + Subclass.asm EbMode thunk)
'
' Message window strategy: create a MESSAGE-ONLY window from the pre-registered
' "STATIC" class and install our own procedure with SetWindowLongW(GWL_WNDPROC).
' This is deliberate, not stylistic: this VB6 build accepts AddressOf only as a
' DIRECT API ARGUMENT (see WinProbe.bas "EnumWindows AddressOf EnumProc, 0") and
' rejects "p = AddressOf Foo" with a syntax error even in a 12-line module
' (measured in isolation). RegisterClassExW would have required the address as a
' value inside a WNDCLASSEXW structure, which is exactly what is unavailable.
' ---------------------------------------------------------------------------

Public Const WM_SOCKET As Long = &H401          ' WM_USER + 1
Public Const WM_TIMER_MSG As Long = &H113
Public Const FD_READ As Long = &H1
Public Const FD_WRITE As Long = &H2
Public Const FD_ACCEPT As Long = &H8
Public Const FD_CLOSE As Long = &H20
Public Const AF_INET As Long = 2
Public Const SOCK_STREAM As Long = 1
Public Const IPPROTO_TCP As Long = 6
Public Const SOL_SOCKET As Long = &HFFFF
Public Const SO_REUSEADDR As Long = 4
' (contract v1.3.1 H2): the EXCLUSIVE option is ~SO_REUSEADDR = -5.
' With REUSEADDR two IDE processes can both bind 8766 and ownership is undefined.
Public Const SO_EXCLUSIVEADDRUSE As Long = -5
Public Const HWND_MESSAGE As Long = -3
Public Const GWL_WNDPROC As Long = -4
Public Const WSA_INVALID As Long = -1

Public Type WSADATA_BUF
    b(0 To 511) As Byte
End Type

Public Type SOCKADDR_IN
    sin_family As Integer
    sin_port As Integer
    sin_addr As Long
    sin_zero(0 To 7) As Byte
End Type

Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVersionRequested As Integer, ByRef lpWSAData As WSADATA_BUF) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long
Private Declare Function socket Lib "ws2_32.dll" (ByVal af As Long, ByVal stype As Long, ByVal protocol As Long) As Long
Private Declare Function bind Lib "ws2_32.dll" (ByVal s As Long, ByRef name As SOCKADDR_IN, ByVal namelen As Long) As Long
Private Declare Function listen Lib "ws2_32.dll" (ByVal s As Long, ByVal backlog As Long) As Long
Private Declare Function accept Lib "ws2_32.dll" (ByVal s As Long, ByRef addr As SOCKADDR_IN, ByRef addrlen As Long) As Long
Private Declare Function recv Lib "ws2_32.dll" (ByVal s As Long, ByVal bufPtr As Long, ByVal lng As Long, ByVal flags As Long) As Long
Private Declare Function send Lib "ws2_32.dll" (ByVal s As Long, ByVal bufPtr As Long, ByVal lng As Long, ByVal flags As Long) As Long
Private Declare Function closesocket Lib "ws2_32.dll" (ByVal s As Long) As Long
Private Declare Function setsockopt Lib "ws2_32.dll" (ByVal s As Long, ByVal level As Long, ByVal optname As Long, ByRef optval As Any, ByVal optlen As Long) As Long
Private Declare Function WSAAsyncSelect Lib "ws2_32.dll" (ByVal s As Long, ByVal hWnd As Long, ByVal wMsg As Long, ByVal lEvent As Long) As Long
Private Declare Function htons Lib "ws2_32.dll" (ByVal hostshort As Integer) As Integer
Private Declare Function WSAGetLastError Lib "ws2_32.dll" () As Long

Private Declare Function CreateWindowExW Lib "user32" (ByVal dwExStyle As Long, ByVal lpClassName As Long, ByVal lpWindowName As Long, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long
Private Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function DefWindowProcW Lib "user32" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal uIDEvent As Long) As Long

Private m_sink As Object
Private m_hWnd As Long

' --- lifecycle -------------------------------------------------------------

Public Function WsaStart() As Boolean
    Dim d As WSADATA_BUF
    WsaStart = (WSAStartup(&H202, d) = 0)
End Function

Public Sub WsaStop()
    On Error Resume Next
    WSACleanup
End Sub

Public Sub WsaSetSink(o As Object)
    Set m_sink = o
End Sub

' The window procedure. Must be a Public procedure in a standard module.
Public Function WsaWndProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    If uMsg = WM_SOCKET Or uMsg = WM_TIMER_MSG Then
        If Not m_sink Is Nothing Then
            WsaWndProc = m_sink.OnWsaMessage(uMsg, wParam, lParam)
            Exit Function
        End If
    End If
    WsaWndProc = DefWindowProcW(hWnd, uMsg, wParam, lParam)
End Function

Public Function WsaMessageWindow() As Long
    Dim clsName As String
    Dim wndName As String
    Dim hInst As Long
    If m_hWnd <> 0 Then
        WsaMessageWindow = m_hWnd
        Exit Function
    End If
    clsName = "STATIC"
    wndName = "VB6MCPWSA"
    hInst = App.hInstance
    m_hWnd = CreateWindowExW(0, StrPtr(clsName), StrPtr(wndName), 0, 0, 0, 0, 0, HWND_MESSAGE, 0, hInst, 0)
    If m_hWnd = 0 Then
        WsaMessageWindow = 0
        Exit Function
    End If
    ' AddressOf must stay a direct API argument (see header note).
    SetWindowLongW m_hWnd, GWL_WNDPROC, AddressOf WsaWndProc
    WsaMessageWindow = m_hWnd
End Function

Public Sub WsaDestroyWindow()
    On Error Resume Next
    If m_hWnd <> 0 Then
        DestroyWindow m_hWnd
        m_hWnd = 0
    End If
End Sub

' --- sockets ---------------------------------------------------------------

Public Function WsaListenSocket(ByVal port As Long, ByRef outErr As Long) As Long
    Dim s As Long
    Dim a As SOCKADDR_IN
    Dim one As Long
    outErr = 0
    s = socket(AF_INET, SOCK_STREAM, IPPROTO_TCP)
    If s = WSA_INVALID Then
        outErr = WSAGetLastError()
        WsaListenSocket = WSA_INVALID
        Exit Function
    End If
    ' SO_EXCLUSIVEADDRUSE, NOT SO_REUSEADDR. With REUSEADDR a second
    ' IDE could also bind 8766 and the port ownership was UNDEFINED (the hub had to
    ' guess). With EXCLUSIVEADDRUSE the second bind fails with WSAEADDRINUSE (10048).
    one = 1
    setsockopt s, SOL_SOCKET, SO_EXCLUSIVEADDRUSE, one, 4
    a.sin_family = AF_INET
    a.sin_port = htons(CLngToInt(port))
    a.sin_addr = 0
    If bind(s, a, LenB(a)) <> 0 Then
        outErr = WSAGetLastError()
        closesocket s
        WsaListenSocket = WSA_INVALID
        Exit Function
    End If
    If listen(s, 32) <> 0 Then
        outErr = WSAGetLastError()
        closesocket s
        WsaListenSocket = WSA_INVALID
        Exit Function
    End If
    WsaListenSocket = s
End Function

Private Function CLngToInt(ByVal v As Long) As Integer
    If v > 32767 Then
        CLngToInt = CInt(v - 65536)
    Else
        CLngToInt = CInt(v)
    End If
End Function

Public Function WsaRegister(ByVal s As Long, ByVal events As Long) As Boolean
    If m_hWnd = 0 Then Exit Function
    WsaRegister = (WSAAsyncSelect(s, m_hWnd, WM_SOCKET, events) = 0)
End Function

Public Function WsaAccept(ByVal listenSock As Long) As Long
    Dim a As SOCKADDR_IN
    Dim l As Long
    l = LenB(a)
    WsaAccept = accept(listenSock, a, l)
End Function

Public Function WsaRecv(ByVal s As Long, ByVal bufPtr As Long, ByVal lng As Long) As Long
    WsaRecv = recv(s, bufPtr, lng, 0)
End Function

Public Function WsaSend(ByVal s As Long, ByVal bufPtr As Long, ByVal lng As Long) As Long
    WsaSend = send(s, bufPtr, lng, 0)
End Function

Public Sub WsaClose(ByVal s As Long)
    On Error Resume Next
    If s <> WSA_INVALID And s <> 0 Then closesocket s
End Sub

Public Function WsaLastError() As Long
    WsaLastError = WSAGetLastError()
End Function

' --- timer -----------------------------------------------------------------

Public Function WsaSetTimer(ByVal ms As Long) As Boolean
    If m_hWnd = 0 Then Exit Function
    WsaSetTimer = (SetTimer(m_hWnd, 1, ms, 0) <> 0)
End Function

Public Sub WsaKillTimer()
    On Error Resume Next
    If m_hWnd <> 0 Then KillTimer m_hWnd, 1
End Sub
