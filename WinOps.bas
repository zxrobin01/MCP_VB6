Attribute VB_Name = "WinOps"
Option Explicit

' =====================================================================
' TASK-P1-069 (contract v1.5.0): message-only window primitives.
'   read_active_dialogs   - enumerate THIS process's active dialogs + buttons
'   close_active_dialogs  - click a SAFE button of an active dialog
'   send_window_message   - SendMessage/PostMessage to a hwnd of THIS process
' TASK-P1-067 (opt-in, DEFAULT OFF): code-pane mouse wheel by SUBCLASSING the
'   code-pane window (class "VbaWindow"). NO keyboard hook, NO global hook,
'   NO WH_MOUSE at all: the whole feature is SetWindowLong/CallWindowProc on
'   windows of our own process.
'
' MEMORY DISCIPLINE (auditable by grep - see TASK-P1-069 acceptance (5)): the
' ONLY Win32 entry points in this module are the message/enumeration/geometry
' family, the subclassing trio and (TASK-P1-079) the advapi32 REGISTRY family
' (RegOpenKeyExA/RegCreateKeyExA/RegSetValueExA/RegQueryValueExA/RegCloseKey). There is NO WriteProcessMemory, NO
' ReadProcessMemory, NO VirtualAlloc, NO CreateRemoteThread, NO SetWindowsHookEx
' and NO IAT patching anywhere here. "close a dialog" means "send a button
' message", never "write into another process".
' =====================================================================

Private Type WO_POINT
    x As Long
    y As Long
End Type

Private Type WO_RECT
    l As Long
    t As Long
    r As Long
    b As Long
End Type

' ---- message / enumeration / geometry -------------------------------------
Private Declare Function WoSendMessageW Lib "user32" Alias "SendMessageW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoPostMessageW Lib "user32" Alias "PostMessageW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoEnumWindows Lib "user32" Alias "EnumWindows" (ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function WoEnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal h As Long, ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function WoGetWindowTextW Lib "user32" Alias "GetWindowTextW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function WoGetClassNameW Lib "user32" Alias "GetClassNameW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function WoGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal h As Long, ByRef rc As WO_RECT) As Long
Private Declare Function WoGetParent Lib "user32" Alias "GetParent" (ByVal h As Long) As Long
Private Declare Function WoGetWindow Lib "user32" Alias "GetWindow" (ByVal h As Long, ByVal cmd As Long) As Long
Private Declare Function WoIsWindow Lib "user32" Alias "IsWindow" (ByVal h As Long) As Long
Private Declare Function WoIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal h As Long) As Long
Private Declare Function WoIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal h As Long) As Long
Private Declare Function WoGetDlgCtrlID Lib "user32" Alias "GetDlgCtrlID" (ByVal h As Long) As Long
Private Declare Function WoGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
Private Declare Function WoGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal h As Long, ByRef pid As Long) As Long
Private Declare Function WoGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function WoGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

' ---- TASK-P1-079: registry access for the settings path --------------------
' WScript.Shell (COM) RegWrite FAILED inside this add-in process - measured with a diagnostic
' probe: RegPut returned False, the read-back right after the write was empty, and
' HKCU\Software\VB6MCP never appeared at all. The settings chain therefore talks to advapi32
' directly, the same library and the same "no COM, no external exe" discipline as Sha256.bas.
' Reads use RegOpenKeyExA (a read must NOT create the key, or "source: default" would turn into
' "source: HKCU" just by opening the dialog).
Private Declare Function WoRegOpenKeyExA Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function WoRegCreateKeyExA Lib "advapi32" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, ByRef phkResult As Long, ByRef lpdwDisposition As Long) As Long
Private Declare Function WoRegSetValueExA Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, ByRef lpData As Any, ByVal cbData As Long) As Long
Private Declare Function WoRegQueryValueExA Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Long) As Long
Private Declare Function WoRegCloseKey Lib "advapi32" Alias "RegCloseKey" (ByVal hKey As Long) As Long

Private Const WO_HKEY_CURRENT_USER As Long = &H80000001
Private Const WO_KEY_QUERY_VALUE As Long = &H1
Private Const WO_KEY_ALL_ACCESS As Long = &HF003F
Private Const WO_REG_SZ As Long = 1



' ---- subclassing trio (TASK-P1-067 only) ----------------------------------
Private Declare Function WoSetWindowLongW Lib "user32" Alias "SetWindowLongW" (ByVal h As Long, ByVal idx As Long, ByVal v As Long) As Long
Private Declare Function WoGetWindowLongW Lib "user32" Alias "GetWindowLongW" (ByVal h As Long, ByVal idx As Long) As Long
Private Declare Function WoCallWindowProcW Lib "user32" Alias "CallWindowProcW" (ByVal prev As Long, ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoDefWindowProcW Lib "user32" Alias "DefWindowProcW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef pt As WO_POINT) As Long
Private Declare Function WoGetKeyState Lib "user32" Alias "GetKeyState" (ByVal vk As Long) As Integer
' TASK-P1-088: used only to read the code pane's first visible line (see WheelPaneTop).
Private Declare Function WoFindWindowExW Lib "user32" Alias "FindWindowExW" (ByVal parent As Long, ByVal after As Long, ByVal cls As Long, ByVal title As Long) As Long

' ---- window messages ------------------------------------------------------
Private Const WM_NULL As Long = &H0
Private Const WM_CLOSE As Long = &H10
Private Const WM_COMMAND As Long = &H111
Private Const WM_KEYDOWN As Long = &H100
Private Const WM_KEYUP As Long = &H101
Private Const WM_CHAR As Long = &H102
Private Const WM_SYSCOMMAND As Long = &H112
Private Const WM_VSCROLL As Long = &H115
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const BM_CLICK As Long = &HF5
Private Const BM_SETSTATE As Long = &HF3
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const EM_LINESCROLL As Long = &HB6
Private Const EM_GETFIRSTVISIBLELINE As Long = &HCE

Private Const SB_LINEUP As Long = 0
Private Const SB_LINEDOWN As Long = 1
Private Const SB_PAGEUP As Long = 2
Private Const SB_PAGEDOWN As Long = 3

Private Const GW_OWNER As Long = 4
Private Const GWLP_WNDPROC As Long = -4
Private Const VK_SHIFT As Long = &H10
Private Const VK_CONTROL As Long = &H11

Private Const WO_MAX As Long = 1200
Private Const WO_KIDS_MAX As Long = 80
Private Const WHEEL_MAX As Long = 32

Private m_woHwnd(0 To WO_MAX - 1) As Long
Private m_woCount As Long

' ---- TASK-P1-067 wheel state (module-level, so it MUST live here: VB6 puts
' every module-level declaration before the first procedure) ----
Private m_whHwnd(0 To WHEEL_MAX - 1) As Long
Private m_whPrev(0 To WHEEL_MAX - 1) As Long
Private m_whCount As Long
Private m_whSelf As Long
Private m_whEnabled As Boolean
Private m_whInstalled As Boolean
Private m_whLastScan As Long
Private m_whLines As Long
Private m_whEvents As Long
Private m_whHandled As Long
Private m_whForwarded As Long
Private m_whScans As Long
Private m_whInProc As Boolean
' TASK-P1-130: TRUE only after this add-in has ONCE positively identified a code pane from the
' VBE caption table. It ARMS the empty-table fallback in WheelConsider: at add-in start-up (no
' project, no code pane at all) the table is legitimately empty and hooking anything there would
' gain nothing while making a non-code window the first entry of the receipt's target fields.
Private m_whEverVerified As Boolean
' TASK-P1-088 observability: the last WM_MOUSEWHEEL we saw, what it carried, what we did
' with it, and the pane's first visible line before/after our scroll messages. All of it
' is READ-ONLY state surfaced through WheelStateObj (mcp_info / ide_state).
Private m_whLastWparam As Long
Private m_whLastLparam As Long
Private m_whLastHwnd As Long
Private m_whLastPane As Long
Private m_whLastDelta As Long
Private m_whLastScroll As Long
Private m_whLastTopBefore As Long
Private m_whLastTopAfter As Long
Private m_whLastTitle As String
Private m_whLastReason As String

' ---- TASK-P1-067 final fix: move the pane through the VBIDE object model ----
' WHY. Measured 2026-10-09 on the REAL code pane (hwnd=23464986, title
' "P1067 - modLong (Code)"): WM_VSCROLL + SB_LINEUP/SB_LINEDOWN/SB_PAGEUP/
' SB_PAGEDOWN 5/5 left CodePane.TopLine at 1 even though SendMessage returned
' and the message reached our own WndProc. The pane does not respond to the
' scroll message, so the wheel now sets CodePane.TopLine in-process instead.
' HOW. This is a STANDARD module: it has no VBIDE reference. CodeTools (the only
' holder of the VBE) is wired in by McpServer as a late-bound sink, so the two
' helpers below can be called without a new type library.
Private m_whSink As Object
' Caption of every hooked pane, in lockstep with m_whHwnd/m_whPrev. The caption
' IS the hwnd -> CodePane key: measured GetWindowText == CodePane.Window.Caption
' verbatim (the "(Code)" suffix is never parsed: localisation changes it).
Private m_whTitle(0 To WHEEL_MAX - 1) As String
' Windows of class VbaWindow that are NOT code panes (the Immediate window has
' the same class). They are skipped, never subclassed, and listed in the receipt
' so the observable fields can no longer point at the wrong window.
Private Const WHEEL_SKIP_MAX As Long = 16
Private m_whSkipHwnd(0 To WHEEL_SKIP_MAX - 1) As Long
Private m_whSkipCls(0 To WHEEL_SKIP_MAX - 1) As String
Private m_whSkipTitle(0 To WHEEL_SKIP_MAX - 1) As String
Private m_whSkipReason(0 To WHEEL_SKIP_MAX - 1) As String
Private m_whSkipCount As Long
' The caption table built by the LAST WheelScan (read-only receipt data).
Private m_whCaps As Object
' TASK-P1-132 (D-130-01): where the caption table came from on the last scan
' ("all" | "active_fallback" | "empty") and how many captions it held. Reported in
' ide_state.wheel so a missing hook has an explained cause, not just a symptom.
Private m_whCapsSource As String
Private m_whCapsCount As Long
' TASK-P1-132 (D-130-01): which channel resolved the last caption -> CodePane lookup
' ("caption" | "component" | "active_name" | "none").
Private m_whResolveVia As String
' TASK-P1-132 (D-130-01): the REAL status/error of the last TopLine read/write, and the
' pending scroll that the WndProc could not perform itself.
Private m_whTopStatus As String
Private m_whTopErr As String
Private m_whPendOn As Boolean
Private m_whPendTitle As String
Private m_whPendLines As Long
Private m_whPendPages As Long
Private m_whPendTries As Long
Private m_whPendResult As String
Private m_whLastPaneCaption As String

' ---- TASK-P1-084: armed, one-shot CONFIRM of a modal WE opened ---------------
' Why this lives here and not in the caller: the IDE's own Add Form modal runs
' its OWN message loop. A tool call that opens that modal cannot click it - its
' DoEvents is trapped INSIDE the modal loop, so the code after DoEvents never
' runs and the click never happens. The click therefore has to come from a
' message the modal loop DISPATCHES: the add-in's watchdog tick (ServerForm's
' tmrWatch -> WinOps.WheelTick) is exactly such a message. Arm here, execute the
' menu command in CodeTools, pump, and let the tick do the click.
Private m_cfmArmed As Boolean
Private m_cfmDeadline As Long
Private m_cfmClicked As Boolean
Private m_cfmDlg As Long
Private m_cfmBtn As Long
Private m_cfmText As String
Private m_cfmTries As Long
Private m_cfmResult As String
Private m_cfmArmedAt As Long
Private m_cfmScan(0 To 31) As Long
Private m_cfmScanN As Long

' =====================================================================
' small helpers
' =====================================================================

' HKCU\<subKey>\<valueName> as a string; "" when the key/value is absent or unreadable.
Public Function RegGetString(ByVal subKey As String, ByVal valueName As String) As String
    On Error Resume Next
    RegGetString = ""
    Dim hKey As Long, rc As Long, typ As Long, n As Long
    Dim buf As String
    hKey = 0
    rc = WoRegOpenKeyExA(WO_HKEY_CURRENT_USER, subKey, 0&, WO_KEY_QUERY_VALUE, hKey)
    If rc <> 0 Then Exit Function
    If hKey = 0 Then Exit Function
    n = 512
    buf = String$(n, 0)
    rc = WoRegQueryValueExA(hKey, valueName, 0&, typ, ByVal buf, n)
    WoRegCloseKey hKey
    If rc <> 0 Then Exit Function
    If n > 1 Then RegGetString = Trim$(Left$(buf, n - 1))
End Function

' Write HKCU\<subKey>\<valueName> = v as REG_SZ. True only when the registry call succeeded.
Public Function RegSetString(ByVal subKey As String, ByVal valueName As String, ByVal v As String) As Boolean
    On Error Resume Next
    RegSetString = False
    Dim hKey As Long, disp As Long, rc As Long, s As String
    hKey = 0
    rc = WoRegCreateKeyExA(WO_HKEY_CURRENT_USER, subKey, 0&, vbNullString, 0&, WO_KEY_ALL_ACCESS, 0&, hKey, disp)
    If rc <> 0 Then Exit Function
    If hKey = 0 Then Exit Function
    s = v
    rc = WoRegSetValueExA(hKey, valueName, 0&, WO_REG_SZ, ByVal s, Len(s) + 1)
    WoRegCloseKey hKey
    RegSetString = (rc = 0)
End Function

' TASK-P1-079: the settings path addresses HKCU\Software\VB6MCP only, so callers pass just the
' leaf name; keeping the path here keeps backslashes out of the form's code.
Public Function McpSettingGet(ByVal leaf As String) As String
    McpSettingGet = RegGetString("Software\VB6MCP", leaf)
End Function

Public Function McpSettingSet(ByVal leaf As String, ByVal v As String) As Boolean
    McpSettingSet = RegSetString("Software\VB6MCP", leaf, v)
End Function

Private Function WoText(ByVal h As Long) As String
    Dim b As String, n As Long
    On Error Resume Next
    b = String$(512, vbNullChar)
    n = WoGetWindowTextW(h, StrPtr(b), 512)
    If n > 0 Then WoText = Left$(b, n) Else WoText = ""
End Function

Private Function WoCls(ByVal h As Long) As String
    Dim b As String, n As Long
    On Error Resume Next
    b = String$(256, vbNullChar)
    n = WoGetClassNameW(h, StrPtr(b), 256)
    If n > 0 Then WoCls = Left$(b, n) Else WoCls = ""
End Function

Private Function WoPidOf(ByVal h As Long) As Long
    Dim p As Long
    p = 0
    WoGetWindowThreadProcessId h, p
    WoPidOf = p
End Function

Private Function WoMine(ByVal h As Long) As Boolean
    If h = 0 Then Exit Function
    If WoIsWindow(h) = 0 Then Exit Function
    WoMine = (WoPidOf(h) = WoGetCurrentProcessId())
End Function

' decimal, 0x hex, &H hex or negative decimal; ok=False when unparsable.
Private Function WoNum(ByVal s As String, ByRef ok As Boolean) As Long
    Dim t As String, neg As Boolean, i As Long, ch As String, d As Long
    Dim v As Double
    ok = False
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    If Left$(t, 1) = "-" Then
        neg = True
        t = Mid$(t, 2)
    End If
    If LCase$(Left$(t, 2)) = "&h" Or LCase$(Left$(t, 2)) = "0x" Then
        t = Mid$(t, 3)
        If Len(t) = 0 Then Exit Function
        v = 0
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If ch >= "0" And ch <= "9" Then
                d = Asc(ch) - 48
            ElseIf LCase$(ch) >= "a" And LCase$(ch) <= "f" Then
                d = Asc(LCase$(ch)) - 87
            Else
                Exit Function
            End If
            v = v * 16 + d
            If v > 4294967295# Then Exit Function
        Next
    Else
        For i = 1 To Len(t)
            ch = Mid$(t, i, 1)
            If ch < "0" Or ch > "9" Then Exit Function
        Next
        v = CDbl(Val(t))
    End If
    If neg Then v = -v
    If v < -2147483648# Or v > 4294967295# Then Exit Function
    If v > 2147483647# Then v = v - 4294967296#
    WoNum = CLng(v)
    ok = True
End Function

Private Function LoWord(ByVal v As Long) As Long
    Dim w As Long
    w = v And &HFFFF&
    If w > 32767 Then w = w - 65536
    LoWord = w
End Function

Private Function HiWord(ByVal v As Long) As Long
    Dim w As Long
    w = (v \ 65536) And &HFFFF&
    If w > 32767 Then w = w - 65536
    HiWord = w
End Function

' =====================================================================
' TASK-P1-069 (1): read_active_dialogs
' Enumerate the ACTIVE dialogs of THIS process (default class #32770) and
' report class / title / geometry / owner / whether the OWNER is DISABLED (the
' only reliable "this box is blocking its parent" signal - "a window exists" is
' not enough) plus the child controls and the clickable button list.
' READ-ONLY: it enumerates and reads text; it never sends anything.
' =====================================================================

Private Function WoCollect(ByVal h As Long, ByVal lp As Long) As Long
    If m_woCount < WO_MAX Then
        m_woHwnd(m_woCount) = h
        m_woCount = m_woCount + 1
    End If
    WoCollect = 1
End Function

Private Sub WoScanTops()
    m_woCount = 0
    WoEnumWindows AddressOf WoCollect, 0&
End Sub

' Every DESCENDANT of h (EnumChildWindows walks the whole subtree).
Private Sub WoScanKids(ByVal h As Long)
    m_woCount = 0
    WoEnumChildWindows h, AddressOf WoCollect, 0&
End Sub

Private Function WoClassMatch(ByVal cls As String, ByVal want As String) As Boolean
    If Len(want) = 0 Then
        WoClassMatch = (LCase$(cls) = "#32770")
        Exit Function
    End If
    If want = "*" Then
        WoClassMatch = True
        Exit Function
    End If
    WoClassMatch = (LCase$(cls) = LCase$(want))
End Function

Private Function WoButtonish(ByVal cls As String) As Boolean
    WoButtonish = (InStr(1, cls, "button", vbTextCompare) > 0)
End Function

Private Sub WoControlEntry(ByVal h As Long, ByRef e As Object)
    Dim rc As WO_RECT
    Set e = CreateObject("Scripting.Dictionary")
    e("hwnd") = h
    e("cls") = WoCls(h)
    e("text") = WoText(h)
    e("id") = 0
    e("visible") = False
    e("enabled") = False
    e("buttonish") = WoButtonish(CStr(e("cls")))
    On Error Resume Next
    e("id") = WoGetDlgCtrlID(h)
    e("visible") = (WoIsWindowVisible(h) <> 0)
    e("enabled") = (WoIsWindowEnabled(h) <> 0)
    rc.l = 0: rc.t = 0: rc.r = 0: rc.b = 0
    WoGetWindowRect h, rc
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("l") = rc.l
    r("t") = rc.t
    r("w") = rc.r - rc.l
    r("h") = rc.b - rc.t
    Set e("rect") = r
    On Error GoTo 0
End Sub

' Array (Collection) of visible matching dialogs of THIS process.
Private Function WoDialogsJson(ByVal cls As String, ByVal wantHwnd As Long, ByVal withChildren As Boolean) As Object
    Dim out As Object
    Set out = New Collection
    Dim tops() As Long, nTops As Long, i As Long
    WoScanTops
    nTops = m_woCount
    If nTops > WO_MAX Then nTops = WO_MAX
    ReDim tops(0 To nTops)
    For i = 0 To nTops - 1
        tops(i) = m_woHwnd(i)
    Next
    Dim pid As Long, fg As Long
    pid = WoGetCurrentProcessId()
    fg = WoGetForegroundWindow()
    For i = 0 To nTops - 1
        Dim h As Long
        h = tops(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                Dim c As String
                c = WoCls(h)
                Dim take As Boolean
                take = WoClassMatch(c, cls)
                If wantHwnd <> 0 Then take = (h = wantHwnd)
                If take Then
                    If WoIsWindowVisible(h) <> 0 Then
                        Dim d As Object
                        Set d = CreateObject("Scripting.Dictionary")
                        d("hwnd") = h
                        d("cls") = c
                        d("title") = WoText(h)
                        d("visible") = True
                        d("enabled") = (WoIsWindowEnabled(h) <> 0)
                        d("foreground") = (h = fg)
                        Dim own As Long
                        own = WoGetWindow(h, GW_OWNER)
                        d("owner") = own
                        d("ownerCls") = ""
                        If own <> 0 Then d("ownerCls") = WoCls(own)
                        d("ownerTitle") = ""
                        If own <> 0 Then d("ownerTitle") = WoText(own)
                        d("ownerDisabled") = False
                        If own <> 0 Then d("ownerDisabled") = (WoIsWindowEnabled(own) = 0)
                        Dim rc As WO_RECT
                        rc.l = 0: rc.t = 0: rc.r = 0: rc.b = 0
                        WoGetWindowRect h, rc
                        Dim r As Object
                        Set r = CreateObject("Scripting.Dictionary")
                        r("l") = rc.l
                        r("t") = rc.t
                        r("w") = rc.r - rc.l
                        r("h") = rc.b - rc.t
                        Set d("rect") = r
                        If withChildren Then
                            Dim kids As Object, btns As Object
                            Set kids = New Collection
                            Set btns = New Collection
                            WoScanKids h
                            Dim kn As Long, j As Long
                            kn = m_woCount
                            If kn > WO_KIDS_MAX Then kn = WO_KIDS_MAX
                            For j = 0 To kn - 1
                                Dim ke As Object
                                WoControlEntry m_woHwnd(j), ke
                                kids.Add ke
                                If CBool(ke("buttonish")) Then btns.Add ke
                            Next
                            d("childCount") = kids.Count
                            Set d("buttons") = btns
                            Set d("controls") = kids
                        End If
                        out.Add d
                    End If
                End If
            End If
        End If
    Next
    Set WoDialogsJson = out
End Function

Public Function ReadActiveDialogsJson(ByVal cls As String, ByVal withChildren As Boolean, ByVal maxN As Long) As String
    On Error Resume Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("ok") = True
    root("pid") = WoGetCurrentProcessId()
    If Len(cls) = 0 Then cls = "#32770"
    root("clsFilter") = cls
    root("readOnly") = True
    Dim ds As Object
    Set ds = WoDialogsJson(cls, 0, withChildren)
    Dim arr As Object
    Set arr = New Collection
    Dim i As Long
    For i = 1 To ds.Count
        If maxN > 0 Then
            If arr.Count >= maxN Then Exit For
        End If
        arr.Add ds.Item(i)
    Next
    root("count") = arr.Count
    Set root("dialogs") = arr
    If arr.Count = 0 Then
        root("note") = "no visible top-level window of this process matches cls=" & cls & ". A box that is up appears here; ownerDisabled=true on it is the signal that it owns (blocks) the IDE"
    End If
    ReadActiveDialogsJson = JsonStringify(root)
End Function

' =====================================================================
' TASK-P1-069 (2): close_active_dialogs
' Explicitly invoked ONLY (nothing in the product calls it automatically).
' Sends exactly one message - BM_CLICK - to a button whose text is on the SAFE
' whitelist and NOT on the danger denylist.
' =====================================================================

Private Function WoTokenHit(ByVal table As String, ByVal norm As String) As Boolean
    Dim p As Long, k As Long, tok As String
    k = 1
    Do While k <= Len(table)
        p = InStr(k, table, "|")
        If p = 0 Then p = Len(table) + 1
        tok = Mid$(table, k, p - k)
        If Len(tok) > 0 Then
            If InStr(1, norm, UCase$(tok), vbTextCompare) > 0 Then
                WoTokenHit = True
                Exit Function
            End If
        End If
        k = p + 1
    Loop
End Function

Private Function WoExactHit(ByVal table As String, ByVal norm As String) As Boolean
    Dim p As Long, k As Long, tok As String
    k = 1
    Do While k <= Len(table)
        p = InStr(k, table, "|")
        If p = 0 Then p = Len(table) + 1
        tok = Mid$(table, k, p - k)
        If Len(tok) > 0 Then
            If StrComp(norm, UCase$(tok), vbTextCompare) = 0 Then
                WoExactHit = True
                Exit Function
            End If
        End If
        k = p + 1
    Loop
End Function

' Whitelist per the card: OK / Yes / Cancel and localized twins. "No" is
' deliberately absent (for a "save changes?" box it discards work). A deny
' token always wins, even when the caption also looks safe.
' TASK-P1-125 (S point 3): cancelling cannot START the host-modal operation, so a
' Cancel/Close/Ignore caption stays allowed on a box the host-modal guard would
' otherwise refuse to confirm. This is what lets make_project clean up a leaked
' Make box safely (Cancel, never OK).
Private Function WoButtonCancelLike(ByVal norm As String) As Boolean
    WoButtonCancelLike = WoExactHit("CANCEL|CLOSE|IGNORE|" & ChrW$(&H53D6) & ChrW$(&H6D88) & "|" & ChrW$(&H5173) & ChrW$(&H95ED) & "|" & ChrW$(&H5FFD) & ChrW$(&H7565), norm)
End Function

Private Function WoButtonAllow(ByVal text As String, ByRef norm As String) As Boolean
    Dim t As String, p As Long
    t = Trim$(text)
    t = Replace(t, "&", "")
    p = InStr(t, "(")
    If p > 0 Then t = Trim$(Left$(t, p - 1))
    Do While Len(t) > 0
        If Right$(t, 1) = "." Then t = Left$(t, Len(t) - 1) Else Exit Do
    Loop
    t = Trim$(t)
    norm = UCase$(t)
    If Len(norm) = 0 Then Exit Function
    Dim bad As String
    bad = "DELETE|REMOVE|UNINSTALL|FORMAT|ERASE|PURGE|DESTROY|OVERWRITE|DISCARD|RESET|" & _
          ChrW$(&H5220) & ChrW$(&H9664) & "|" & ChrW$(&H79FB) & ChrW$(&H9664) & "|" & _
          ChrW$(&H5378) & ChrW$(&H8F7D) & "|" & ChrW$(&H6E05) & ChrW$(&H7A7A) & "|" & _
          ChrW$(&H91CD) & ChrW$(&H7F6E) & "|" & ChrW$(&H8986) & ChrW$(&H76D6) & "|" & _
          ChrW$(&H4E22) & ChrW$(&H5F03) & "|" & ChrW$(&H683C) & ChrW$(&H5F0F) & ChrW$(&H5316)
    If WoTokenHit(bad, norm) Then Exit Function
    Dim good As String
    good = "OK|OKAY|YES|CANCEL|CLOSE|CONTINUE|IGNORE|" & _
           ChrW$(&H786E) & ChrW$(&H5B9A) & "|" & ChrW$(&H597D) & ChrW$(&H7684) & "|" & _
           ChrW$(&H662F) & "|" & ChrW$(&H53D6) & ChrW$(&H6D88) & "|" & _
           ChrW$(&H5173) & ChrW$(&H95ED) & "|" & ChrW$(&H7EE7) & ChrW$(&H7EED) & "|" & _
           ChrW$(&H5FFD) & ChrW$(&H7565)
    WoButtonAllow = WoExactHit(good, norm)
End Function

' TASK-P1-092: "No" is NOT on the default whitelist (for a "save changes?" box it
' discards work) but it is never DESTRUCTIVE, so an EXPLICIT button="No" is the escape
' hatch the body guard suggests. Only the exact captions No / Chinese No qualify.
Private Function WoButtonExplicitNo(ByVal norm As String) As Boolean
    WoButtonExplicitNo = WoExactHit("NO|" & ChrW$(&H5426), norm)
End Function

' TASK-P1-092: destructive-action patterns read from the dialog BODY (title +
' every child control's text). Deliberately broader than the button denylist: a
' destructive confirmation usually puts the verb in the STATIC text, not on the
' button - the add-in removal box is "cannot load X ... remove it from the list"
' with buttons Yes/No/Help only, so the caption alone cannot classify it.
Private Function WoBodyDangerous(ByVal body As String, ByRef hit As String) As Boolean
    hit = ""
    If Len(body) = 0 Then Exit Function
    Dim norm As String
    norm = UCase$(body)
    Dim bad As String
    bad = "REMOVE|DELETE|UNINSTALL|UNREGISTER|OVERWRITE|DISCARD|ERASE|PURGE|DESTROY|RESET|FORMAT|" & _
          ChrW$(&H5220) & ChrW$(&H9664) & "|" & ChrW$(&H79FB) & ChrW$(&H9664) & "|" & _
          ChrW$(&H5378) & ChrW$(&H8F7D) & "|" & ChrW$(&H6CE8) & ChrW$(&H9500) & "|" & _
          ChrW$(&H6E05) & ChrW$(&H9664) & "|" & ChrW$(&H6E05) & ChrW$(&H7A7A) & "|" & _
          ChrW$(&H91CD) & ChrW$(&H7F6E) & "|" & ChrW$(&H8986) & ChrW$(&H76D6) & "|" & _
          ChrW$(&H4E22) & ChrW$(&H5F03) & "|" & ChrW$(&H4E0D) & ChrW$(&H518D) & ChrW$(&H52A0) & ChrW$(&H8F7D)
    Dim p As Long, k As Long, tok As String
    k = 1
    Do While k <= Len(bad)
        p = InStr(k, bad, "|")
        If p = 0 Then p = Len(bad) + 1
        tok = Mid$(bad, k, p - k)
        If Len(tok) > 0 Then
            If InStr(1, norm, UCase$(tok), vbTextCompare) > 0 Then
                hit = tok
                WoBodyDangerous = True
                Exit Function
            End If
        End If
        k = p + 1
    Loop
End Function

' TASK-P1-092: the dialog BODY as the guard sees it = title + every child
' control's text, joined by newlines.
Private Function WoDialogBodyDangerous(ByVal d As Object, ByRef hit As String) As Boolean
    Dim s As String
    s = CStr(d("title"))
    On Error Resume Next
    Dim ctrls As Object
    Set ctrls = d("controls")
    If Not ctrls Is Nothing Then
        Dim i As Long, ce As Object
        For i = 1 To ctrls.Count
            Set ce = ctrls.Item(i)
            s = s & vbLf & CStr(ce("text"))
        Next
    End If
    On Error GoTo 0
    WoDialogBodyDangerous = WoBodyDangerous(s, hit)
End Function
' =====================================================================
' TASK-P1-125 (MCP-01): the dialog-interlock classifier.
'
' Clicking a button whose dialog starts a LONG host-modal operation (VB6's own
' "Make Project" box, or a file chooser that leads into a compile) leaves the host
' inside a modal loop that runs its OWN message pump. This add-in is single
' threaded, so the tool call that sent the click never returns and every later
' WRITE tool is answered reason=busy for as long as the host stays modal (the user
' measured busyMs growing past 43 s with busyTool=close_active_dialogs, recovering
' only by killing the IDE). The risk is therefore CLASSIFIED BEFORE anything is
' sent, and reported on EVERY path - the dry_run preview included, because the
' preview is exactly where the user was misled into clicking.
'
'   none     = a plain message box (every child is a button/static): dismissing it
'              returns control immediately.
'   possible = a custom box (has an edit/combo/list control, or many children):
'              a button may open further host UI.
'   known    = the box IS VB6's Make Project dialog.
Private Function WoHostModalRiskOfDialog(ByVal d As Object, ByRef risk As String, _
                                         ByRef reason As String, ByRef riskClass As String) As Boolean
    risk = "none"
    reason = ""
    riskClass = ""
    Dim title As String
    title = UCase$(CStr(d("title")))
    Dim body As String
    body = title
    Dim hasEdit As Boolean
    Dim nKids As Long
    hasEdit = False
    nKids = 0
    On Error Resume Next
    Dim ctrls As Object
    Set ctrls = d("controls")
    If Not ctrls Is Nothing Then
        Dim i As Long, ce As Object, cc As String
        nKids = ctrls.Count
        For i = 1 To ctrls.Count
            Set ce = ctrls.Item(i)
            cc = LCase$(CStr(ce("cls")))
            body = body & vbLf & UCase$(CStr(ce("text")))
            If InStr(cc, "edit") > 0 Then hasEdit = True
            If InStr(cc, "combo") > 0 Then hasEdit = True
            If InStr(cc, "list") > 0 Then hasEdit = True
        Next
    End If
    On Error GoTo 0
    Dim c As String
    c = LCase$(CStr(d("cls")))
    ' (1) known: VB6's Make Project box. Three independent, language-safe signals.
    If InStr(title, ChrW$(&H751F) & ChrW$(&H6210) & ChrW$(&H5DE5) & ChrW$(&H7A0B)) > 0 Then
        risk = "known"
        riskClass = "make_project_dialog"
        reason = "the dialog title is VB6's own Make Project box; confirming its default button starts a compile INSIDE the host's modal loop, which this single-threaded add-in cannot interrupt"
        WoHostModalRiskOfDialog = True
        Exit Function
    End If
    If InStr(title, "MAKE PROJECT") > 0 Then
        risk = "known"
        riskClass = "make_project_dialog"
        reason = "the dialog title is VB6's own Make Project box; confirming its default button starts a compile INSIDE the host's modal loop, which this single-threaded add-in cannot interrupt"
        WoHostModalRiskOfDialog = True
        Exit Function
    End If
    If InStr(body, "SAVE AS &TYPE:") > 0 Or InStr(body, "SAVE &IN:") > 0 Or InStr(body, "FILE &NAME:") > 0 Then
        risk = "known"
        riskClass = "make_project_dialog"
        reason = "the dialog body carries the Make Project target-selector controls (Save as &type: / Save &in: / File &name:), which only that box has; confirming it starts a compile in the host's modal loop"
        WoHostModalRiskOfDialog = True
        Exit Function
    End If
    ' (2) possible: anything that is NOT a plain message box.
    If c <> "#32770" Then
        risk = "possible"
        riskClass = "custom_window"
        reason = "the window class is " & c & ", not a #32770 message box, so its buttons may open further host UI"
        Exit Function
    End If
    If hasEdit Or nKids > 8 Then
        risk = "possible"
        riskClass = "custom_dialog"
        reason = "the box is a custom dialog (an edit/combo/list control is present, or it has " & CStr(nKids) & " child controls), so a button may open a longer host operation"
        Exit Function
    End If
End Function

' TASK-P1-125: public, read-only. The worst host-modal risk among the dialogs now
' up, plus per-dialog detail as JSON. Returns the dialog count. Used by
' make_project before it decides whether to confirm the Make box.
Public Function HostModalRiskProbe(ByVal cls As String, ByRef risk As String, _
                                   ByRef reason As String, ByRef riskClass As String, _
                                   ByRef detailsJson As String) As Long
    On Error Resume Next
    risk = "none"
    reason = ""
    riskClass = ""
    detailsJson = "[]"
    If Len(cls) = 0 Then cls = "#32770"
    Dim ds As Object
    Set ds = WoDialogsJson(cls, 0, True)
    Dim det As Object
    Set det = New Collection
    Dim n As Long, worst As Long, i As Long
    n = 0
    worst = 0
    If Not ds Is Nothing Then n = ds.Count
    For i = 1 To n
        Dim d As Object
        Set d = ds.Item(i)
        Dim r As String, rs As String, rc As String
        WoHostModalRiskOfDialog d, r, rs, rc
        Dim lvl As Long
        lvl = 0
        If r = "possible" Then lvl = 1
        If r = "known" Then lvl = 2
        Dim e As Object
        Set e = CreateObject("Scripting.Dictionary")
        e("hwnd") = d("hwnd")
        e("title") = d("title")
        e("ownerDisabled") = d("ownerDisabled")
        e("risk") = r
        e("riskClass") = rc
        e("riskReason") = rs
        e("makeDialog") = (rc = "make_project_dialog")
        det.Add e
        If lvl > worst Then
            worst = lvl
            risk = r
            reason = rs
            riskClass = rc
        End If
    Next
    HostModalRiskProbe = n
    detailsJson = JsonStringify(det)
End Function

' TASK-P1-125 (MCP-18): the visible text of every active dialog of this process,
' flattened (title + every child caption). The in-IDE Make writes NO compiler log,
' so a VB6 compile-error box is the only place its text can be read from.
Public Function DialogBodiesText(ByVal cls As String) As String
    On Error Resume Next
    DialogBodiesText = ""
    If Len(cls) = 0 Then cls = "#32770"
    Dim ds As Object
    Set ds = WoDialogsJson(cls, 0, True)
    If ds Is Nothing Then Exit Function
    Dim s As String, i As Long, j As Long
    For i = 1 To ds.Count
        Dim d As Object
        Set d = ds.Item(i)
        s = s & CStr(d("title")) & vbLf
        Dim ctrls As Object
        Set ctrls = d("controls")
        If Not ctrls Is Nothing Then
            For j = 1 To ctrls.Count
                s = s & CStr(ctrls.Item(j)("text")) & vbLf
            Next
        End If
        s = s & vbLf
    Next
    DialogBodiesText = s
End Function

' TASK-P1-125 (MCP-18): best-effort text of a VB6 compile-error box. The box body
' carries the MESSAGE (measured: "Compile error: Sub or Function not defined",
' Chinese: Bianyi Cuowu ...) but not the file/line, so only the text is returned;
' the location comes from the compiler log on the route=cli path. Honest: returns
' False when no compile-error signature is present rather than inventing one.
Public Function CompileErrorTextFromDialogs(ByRef text As String) As Boolean
    On Error Resume Next
    text = ""
    Dim ds As Object
    Set ds = WoDialogsJson("#32770", 0, True)
    If ds Is Nothing Then Exit Function
    Dim ceEn As String, ceZh As String
    ceEn = "COMPILE ERROR"
    ceZh = ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H9519) & ChrW$(&H8BEF)
    Dim i As Long, j As Long
    For i = 1 To ds.Count
        Dim d As Object
        Set d = ds.Item(i)
        Dim s As String
        s = CStr(d("title")) & vbLf
        Dim ctrls As Object
        Set ctrls = d("controls")
        If Not ctrls Is Nothing Then
            For j = 1 To ctrls.Count
                s = s & CStr(ctrls.Item(j)("text")) & vbLf
            Next
        End If
        If InStr(1, s, ceEn, vbTextCompare) > 0 Or InStr(s, ceZh) > 0 Then
            text = Trim$(s)
            CompileErrorTextFromDialogs = True
            Exit Function
        End If
    Next
End Function

Public Function CloseActiveDialogsJson(ByVal hwnd As Long, ByVal onlyButton As String, ByVal dryRun As Boolean, _
                                        Optional ByVal allowHostModal As Boolean = False) As String
    On Error Resume Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("ok") = True
    root("dryRun") = dryRun
    root("explicitCall") = True
    root("allowHostModal") = allowHostModal
    root("policy") = "clicks ONLY buttons whose caption is on the safe whitelist (OK/Yes/Cancel/Close/Continue/Ignore + Chinese twins) and never on the danger denylist; the click is the BM_CLICK window message (a WM_COMMAND fallback is sent only if the button still exists). TASK-P1-125 adds a second, independent guard: a dialog that would start a LONG host-modal operation (VB6's Make Project box) is reported as host_modal_risk=known and is NOT clicked unless allow_host_modal=true."
    ' TASK-P1-125 (discipline 98): the same key set on EVERY path. Without these
    ' defaults a refusal-only key (hostModalGuard / refusedCount / bodyGuard) and the
    ' note vanished from the success path, so a caller could not rely on the shape.
    root("note") = ""
    root("hostModalGuard") = ""
    root("refusedCount") = 0
    root("bodyGuard") = ""
    Dim before As Object
    Set before = WoDialogsJson("", hwnd, True)
    Dim clicked As Object, skipped As Object, riskDialogs As Object
    Set clicked = New Collection
    Set skipped = New Collection
    Set riskDialogs = New Collection
    Dim refusedN As Long, refusedHostN As Long
    Dim worstRisk As String, worstReason As String, worstClass As String
    worstRisk = "none"
    worstReason = ""
    worstClass = ""
    Dim i As Long, j As Long
    Dim want As String
    want = Trim$(Replace(onlyButton, "&", ""))
    For i = 1 To before.Count
        Dim d As Object
        Set d = before.Item(i)
        Dim dlgDangerous As Boolean, bodyHit As String
        dlgDangerous = WoDialogBodyDangerous(d, bodyHit)
        ' ---- TASK-P1-125 (MCP-01): classify BEFORE anything is sent --------------
        Dim dlgRisk As String, dlgRiskReason As String, dlgRiskClass As String
        WoHostModalRiskOfDialog d, dlgRisk, dlgRiskReason, dlgRiskClass
        Dim dlgRiskLvl As Long
        dlgRiskLvl = 0
        If dlgRisk = "possible" Then dlgRiskLvl = 1
        If dlgRisk = "known" Then dlgRiskLvl = 2
        Dim worstLvl As Long
        worstLvl = 0
        If worstRisk = "possible" Then worstLvl = 1
        If worstRisk = "known" Then worstLvl = 2
        If dlgRiskLvl > worstLvl Then
            worstRisk = dlgRisk
            worstReason = dlgRiskReason
            worstClass = dlgRiskClass
        End If
        Dim re As Object
        Set re = CreateObject("Scripting.Dictionary")
        re("hwnd") = d("hwnd")
        re("title") = d("title")
        re("ownerDisabled") = d("ownerDisabled")
        re("risk") = dlgRisk
        re("riskClass") = dlgRiskClass
        re("riskReason") = dlgRiskReason
        riskDialogs.Add re
        Dim btns As Object
        Set btns = btnsOf(d)
        If Not btns Is Nothing Then
            For j = 1 To btns.Count
                Dim b As Object
                Set b = btns.Item(j)
                Dim t As String, norm As String
                t = CStr(b("text"))
                Dim normA As String
                normA = UCase$(Trim$(Replace(t, "&", "")))
                Dim sk As Object
                Set sk = CreateObject("Scripting.Dictionary")
                ' TASK-P1-126 (MCP-14): skipped[] carries the SAME key set on every
                ' refusal branch, so one shape can be read whichever reason fired.
                sk("dialog") = d("hwnd")
                sk("dialogTitle") = d("title")
                sk("hwnd") = b("hwnd")
                sk("text") = t
                sk("id") = b("id")
                sk("normalized") = normA
                sk("method") = "none"
                sk("methodUsed") = "none (refused before any message was sent)"
                sk("clicked") = False
                sk("planned") = False
                sk("sent") = False
                sk("sendResult") = 0
                sk("commandFallback") = 0
                sk("commandFallbackUsed") = False
                sk("stillExists") = (WoIsWindow(CLng(b("hwnd"))) <> 0)
                sk("reason") = ""
                sk("bodyHit") = ""
                sk("advice") = ""
                sk("host_modal_risk") = dlgRisk
                sk("risk_reason") = dlgRiskReason
                sk("risk_class") = dlgRiskClass
                Dim isTarget As Boolean
                isTarget = True
                If Len(want) > 0 Then
                    If StrComp(Trim$(Replace(t, "&", "")), want, vbTextCompare) <> 0 Then isTarget = False
                End If
                If isTarget Then
                    If dlgDangerous And Len(want) = 0 Then
                        sk("reason") = "dangerous_dialog_body"
                        sk("bodyHit") = bodyHit
                        sk("dialogTitle") = d("title")
                        sk("advice") = "refused: the dialog body names a destructive action (" & bodyHit & "); read the box yourself or pass button=""No"" to click No explicitly"
                        skipped.Add sk
                        refusedN = refusedN + 1
                        Exit For
                    ElseIf dlgRisk = "known" And Not allowHostModal And Not WoButtonCancelLike(normA) Then
                        ' TASK-P1-125: this is the click that cost the user the whole
                        ' IDE. Refuse it and say exactly why, on dry_run too.
                        sk("reason") = "host_modal_risk_known"
                        sk("dialogTitle") = d("title")
                        sk("advice") = "refused: confirming this box starts a long host-modal operation (a compile) that this single-threaded add-in cannot interrupt - every later write tool would answer reason=busy (busyTool=close_active_dialogs) until the IDE is restarted. Call again with allow_host_modal=true if you really mean to start it, or use make_project, which handles this box itself."
                        skipped.Add sk
                        refusedHostN = refusedHostN + 1
                        Exit For
                    ElseIf Not CBool(b("enabled")) Then
                        sk("reason") = "button_disabled"
                        skipped.Add sk
                    ElseIf (Not WoButtonAllow(t, norm)) And Not (Len(want) > 0 And WoButtonExplicitNo(norm)) Then
                        sk("reason") = "not_on_safe_whitelist"
                        skipped.Add sk
                    Else
                        Dim ck As Object
                        Set ck = CreateObject("Scripting.Dictionary")
                        ck("dialog") = d("hwnd")
                        ck("dialogTitle") = d("title")
                        ck("hwnd") = b("hwnd")
                        ck("text") = t
                        ck("id") = b("id")
                        ck("normalized") = norm
                        ck("method") = "BM_CLICK"
                        ' TASK-P1-126 (MCP-14): ONE key set on BOTH the dry_run and the
                        ' real-click path. Before this the dry path emitted clicked/planned
                        ' while the real path emitted sendResult/commandFallback, so a
                        ' caller could not compare the two receipts key for key (discipline 98).
                        ck("methodUsed") = "none (dry_run: nothing was sent)"
                        ck("planned") = dryRun
                        ck("clicked") = False
                        ck("sent") = False
                        ck("sendResult") = 0
                        ck("commandFallback") = 0
                        ck("commandFallbackUsed") = False
                        ck("stillExists") = (WoIsWindow(CLng(b("hwnd"))) <> 0)
                        ck("host_modal_risk") = dlgRisk
                        ck("risk_reason") = dlgRiskReason
                        ck("risk_class") = dlgRiskClass
                        ck("allowHostModal") = allowHostModal
                        If dlgDangerous Then ck("dangerousDialogBody") = bodyHit
                        If dryRun Then
                            ck("clicked") = False
                            ck("planned") = True
                            ck("stillExists") = (WoIsWindow(CLng(b("hwnd"))) <> 0)
                        Else
                            ' SendMessage to a window of our own thread runs the
                            ' click synchronously, so "stillExists" is a fact.
                            ck("sendResult") = WoSendMessageW(CLng(b("hwnd")), BM_CLICK, 0, 0)
                            ck("sent") = True
                            ck("methodUsed") = "BM_CLICK"
                            ' BM_CLICK is documented to do nothing on an INACTIVE
                            ' dialog: cross-check, and if the button is still
                            ' there use the dialog's WM_COMMAND/BN_CLICKED path.
                            If WoIsWindow(CLng(b("hwnd"))) <> 0 Then
                                ck("commandFallback") = WoSendMessageW(CLng(d("hwnd")), WM_COMMAND, CLng(b("id")), CLng(b("hwnd")))
                                ck("commandFallbackUsed") = True
                                ck("methodUsed") = "BM_CLICK+WM_COMMAND"
                            End If
                            ck("stillExists") = (WoIsWindow(CLng(b("hwnd"))) <> 0)
                            ck("clicked") = True
                        End If
                        clicked.Add ck
                        ' Exactly ONE button per dialog: clicking every safe
                        ' button in turn would press OK and then Cancel on the
                        ' same box. The first safe, enabled button in the
                        ' dialog's own enumeration order wins.
                        Exit For
                    End If
                End If
            Next
        End If
    Next
    Dim after As Object
    Set after = WoDialogsJson("", hwnd, False)
    Dim remaining As Object
    Set remaining = New Collection
    For i = 1 To after.Count
        remaining.Add after.Item(i)
    Next
    root("dialogsBefore") = before.Count
    root("clickedCount") = clicked.Count
    root("skippedCount") = skipped.Count
    Set root("clicked") = clicked
    Set root("skipped") = skipped
    root("dialogsAfter") = remaining.Count
    Set root("remaining") = remaining
    ' ---- TASK-P1-125 (MCP-01): the host-modal verdict, on EVERY path ---------
    root("host_modal_risk") = worstRisk
    root("host_modal_risk_reason") = worstReason
    root("host_modal_risk_class") = worstClass
    Set root("riskDialogs") = riskDialogs
    root("refusedHostModalCount") = refusedHostN
    If worstRisk = "known" Then
        root("suggested_action") = "Do NOT confirm this box unless you accept that the add-in will stop answering WRITE tools until the operation ends. For VB6's Make Project box use make_project, which pins the target itself so the box never appears, or close_active_dialogs with allow_host_modal=true when you deliberately want the compile to start inside the host."
    ElseIf worstRisk = "possible" Then
        root("suggested_action") = "No known host-modal box is up. A custom dialog is present; if the click does not return quickly, read_active_dialogs (read-only) still answers - see recovery_hint."
    Else
        root("suggested_action") = "No host-modal risk detected; a plain message box is dismissed immediately by the click."
    End If
    root("recovery_hint") = "If a write tool answers reason=busy right after this call, the host is inside a modal loop opened by the button just clicked: (1) WAIT - the write usually finishes on its own; check inflight_since_ms / inflight_tool on the busy receipt; (2) INSPECT WITHOUT WRITING - read_active_dialogs is read-only and is NOT gated by the busy check, so it still answers (ownerDisabled=true identifies the blocking box); (3) if busyMs keeps growing, restart the IDE: end the VB6 process and start it again on the same .vbp - the add-in re-attaches on load and nothing in the project directory is lost, because the click only opened a modal, it did not delete anything."
    If clicked.Count = 0 And before.Count > 0 Then
        root("note") = "no safe, enabled button matched; nothing was clicked"
    End If
    If refusedHostN > 0 Then
        root("hostModalGuard") = "host_modal_risk_known: the box that was up would have started a long host-modal operation, so NOTHING was sent; pass allow_host_modal=true to click it anyway"
    End If
    If refusedN > 0 Then
        root("refusedCount") = refusedN
        root("bodyGuard") = "dangerous_dialog_body: the dialog text named a destructive action, so the DEFAULT click was refused and nothing was sent; pass button=""No"" to click No explicitly, or handle the box yourself"
    End If
    CloseActiveDialogsJson = JsonStringify(root)
End Function

' =====================================================================
' TASK-P1-084: the ARMED, ONE-SHOT confirm used by the IDE-native form route.
' Arm -> (the caller executes the IDE's own "Add Form" menu command) -> the
' watchdog tick clicks that modal's confirm button -> the caller pumps until the
' click lands. Why the tick has to do it: the modal runs its OWN message loop, so
' a tool call that opened it is suspended INSIDE that loop and can never reach
' the code that would click it; the watchdog's WM_TIMER is dispatched BY that
' loop, which is the only place the click can come from.
' SAFETY: windows of THIS process only, VISIBLE #32770 only, an ENABLED button
' whose caption (accelerators stripped) is exactly Open/OK/确定/打开 only, and a
' refusal to arm while a dialog of ours is already up (never click into a box this
' call did not open). No hook, no memory write: one BM_CLICK window message.
' =====================================================================

Private Function WoCfmCollect(ByVal h As Long, ByVal lp As Long) As Long
    On Error Resume Next
    If m_cfmScanN < 32 Then
        If WoIsWindowVisible(h) <> 0 Then
            If WoPidOf(h) = WoGetCurrentProcessId() Then
                m_cfmScan(m_cfmScanN) = h
                m_cfmScanN = m_cfmScanN + 1
            End If
        End If
    End If
    WoCfmCollect = 1
End Function

' Visible top-level windows of this process, compacted down to #32770 dialogs.
Private Function WoCfmDialogs() As Long
    On Error Resume Next
    Dim i As Long, d As Long, n As Long
    m_cfmScanN = 0
    WoEnumWindows AddressOf WoCfmCollect, 0&
    n = 0
    For i = 0 To m_cfmScanN - 1
        d = m_cfmScan(i)
        If LCase$(WoCls(d)) = "#32770" Then
            m_cfmScan(n) = d
            n = n + 1
        End If
    Next
    m_cfmScanN = n
    WoCfmDialogs = n
End Function

' Caption normalized the way the whitelist needs it: accelerators stripped,
' whitespace and trailing dots removed, upper case.
Private Function WoCfmNorm(ByVal t As String) As String
    Dim s As String
    s = Trim$(Replace(t, "&", ""))
    Do While Len(s) > 0
        If Right$(s, 1) = "." Then
            s = Trim$(Left$(s, Len(s) - 1))
        Else
            Exit Do
        End If
    Loop
    WoCfmNorm = UCase$(s)
End Function

' The confirm button of dialog d, if it is one we are willing to press.
' The IDE's Add Form dialog is a file-style dialog: its confirm button is "&Open"
' (localized "打开"), which is deliberately NOT on close_active_dialogs' safe
' whitelist - that tool would skip it and leave the modal over the IDE.
Private Function WoCfmButton(ByVal d As Long, ByRef outText As String) As Long
    On Error Resume Next
    Dim i As Long, h As Long, t As String
    outText = ""
    If d = 0 Then Exit Function
    m_cfmScanN = 0
    WoEnumChildWindows d, AddressOf WoCfmCollect, 0&
    For i = 0 To m_cfmScanN - 1
        h = m_cfmScan(i)
        If WoButtonish(WoCls(h)) Then
            If WoIsWindowEnabled(h) <> 0 Then
                t = WoCfmNorm(WoText(h))
                If t = "OPEN" Or t = "OK" Or t = "OKAY" Then
                    outText = WoText(h)
                    WoCfmButton = h
                    Exit Function
                End If
                If t = ChrW$(&H6253) & ChrW$(&H5F00) Then   ' "打开"
                    outText = WoText(h)
                    WoCfmButton = h
                    Exit Function
                End If
                If t = ChrW$(&H786E) & ChrW$(&H5B9A) Then   ' "确定"
                    outText = WoText(h)
                    WoCfmButton = h
                    Exit Function
                End If
            End If
        End If
    Next
End Function

' Arm the one-shot confirm. FALSE (with WoConfirmStatus explaining why) when a
' visible dialog of THIS process is already up: "the dialog I opened" would be
' ambiguous and we must never press a button in a box we did not open.
Public Function WoConfirmArm(ByVal timeoutMs As Long) As Boolean
    On Error Resume Next
    WoConfirmArm = False
    m_cfmArmed = False
    m_cfmClicked = False
    m_cfmDlg = 0
    m_cfmBtn = 0
    m_cfmText = ""
    m_cfmTries = 0
    m_cfmArmedAt = WoGetTickCount()
    If WoCfmDialogs() > 0 Then
        m_cfmResult = "dialog_already_open"
        Exit Function
    End If
    m_cfmArmed = True
    m_cfmDeadline = m_cfmArmedAt + timeoutMs
    m_cfmResult = "armed"
    WoConfirmArm = True
End Function

Public Sub WoConfirmDisarm()
    m_cfmArmed = False
    If m_cfmResult = "armed" Or m_cfmResult = "waiting" Or m_cfmResult = "waiting_button" Then m_cfmResult = "disarmed"
End Sub

Public Function WoConfirmClicked() As Boolean
    WoConfirmClicked = m_cfmClicked
End Function

Public Function WoConfirmArmed() As Boolean
    WoConfirmArmed = m_cfmArmed
End Function

Public Function WoConfirmStatus() As String
    WoConfirmStatus = m_cfmResult
End Function

Public Function WoConfirmDetail() As String
    WoConfirmDetail = "dlg=" & CStr(m_cfmDlg) & " btn=" & CStr(m_cfmBtn) & " text=" & m_cfmText & _
        " ticks=" & CStr(m_cfmTries) & " ms=" & CStr(WoGetTickCount() - m_cfmArmedAt)
End Function

' Called from the watchdog tick BEFORE the wheel's own early exit, because the
' confirm must work with the wheel OFF (the default). Also reached from the message
' pump, which is harmless: the state machine is idempotent.
Public Sub WoConfirmTick()
    On Error Resume Next
    If Not m_cfmArmed Then Exit Sub
    m_cfmTries = m_cfmTries + 1
    Dim n As Long, d As Long, btn As Long, t As String
    n = WoCfmDialogs()
    If n = 0 Then
        If WoGetTickCount() > m_cfmDeadline Then
            m_cfmArmed = False
            m_cfmResult = "no_dialog"
        Else
            m_cfmResult = "waiting"
        End If
        Exit Sub
    End If
    d = m_cfmScan(0)
    m_cfmDlg = d
    btn = WoCfmButton(d, t)
    If btn <> 0 Then
        m_cfmBtn = btn
        m_cfmText = t
        WoSendMessageW btn, BM_CLICK, 0, 0
        m_cfmClicked = True
        m_cfmArmed = False
        m_cfmResult = "clicked"
        Exit Sub
    End If
    If WoGetTickCount() > m_cfmDeadline Then
        ' Honest bound: no confirm button we are allowed to press was found, so the
        ' box WE opened is dismissed (WM_CLOSE) instead of being left over the IDE.
        ' Only reachable while armed, and arming verified no dialog existed before.
        m_cfmResult = "timeout_no_confirm_button"
        WoSendMessageW d, WM_CLOSE, 0, 0
        m_cfmArmed = False
    Else
        m_cfmResult = "waiting_button"
    End If
End Sub

Private Function btnsOf(ByVal d As Object) As Object
    On Error Resume Next
    Set btnsOf = d("buttons")
End Function

' =====================================================================
' TASK-P1-069 (1): send_window_message
' "click / send a key / send a raw message" to a hwnd of THIS process.
' =====================================================================

Private Function WoMsgByName(ByVal name As String, ByRef ok As Boolean) As Long
    ok = True
    Select Case LCase$(Trim$(name))
        Case "click", "bm_click": WoMsgByName = BM_CLICK
        Case "wm_command", "command": WoMsgByName = WM_COMMAND
        Case "wm_keydown", "keydown": WoMsgByName = WM_KEYDOWN
        Case "wm_keyup", "keyup": WoMsgByName = WM_KEYUP
        Case "wm_char", "char": WoMsgByName = WM_CHAR
        Case "wm_close", "close": WoMsgByName = WM_CLOSE
        Case "wm_mousewheel", "mousewheel", "wheel": WoMsgByName = WM_MOUSEWHEEL
        Case "wm_vscroll", "vscroll": WoMsgByName = WM_VSCROLL
        Case "wm_syscommand": WoMsgByName = WM_SYSCOMMAND
        Case "wm_null": WoMsgByName = WM_NULL
        Case "bm_setstate": WoMsgByName = BM_SETSTATE
        Case "cb_showdropdown": WoMsgByName = CB_SHOWDROPDOWN
        Case "em_linescroll": WoMsgByName = EM_LINESCROLL
        Case Else
            ok = False
    End Select
End Function

Private Function WoVkByName(ByVal name As String, ByRef ok As Boolean) As Long
    Dim s As String, c As String
    s = UCase$(Trim$(name))
    ok = True
    If Len(s) = 0 Then
        ok = False
        Exit Function
    End If
    If Len(s) > 1 Then
        If Left$(s, 2) = "0X" Or Left$(s, 2) = "&H" Then
            WoVkByName = WoNum(s, ok)
            Exit Function
        End If
    End If
    If Len(s) = 1 Then
        c = s
        If c >= "A" And c <= "Z" Then
            WoVkByName = Asc(c)
            Exit Function
        End If
        If c >= "0" And c <= "9" Then
            WoVkByName = Asc(c)
            Exit Function
        End If
    End If
    Select Case s
        Case "ENTER", "RETURN": WoVkByName = &HD
        Case "ESC", "ESCAPE": WoVkByName = &H1B
        Case "TAB": WoVkByName = &H9
        Case "SPACE": WoVkByName = &H20
        Case "BACKSPACE", "BACK": WoVkByName = &H8
        Case "UP": WoVkByName = &H26
        Case "DOWN": WoVkByName = &H28
        Case "LEFT": WoVkByName = &H25
        Case "RIGHT": WoVkByName = &H27
        Case "HOME": WoVkByName = &H24
        Case "END": WoVkByName = &H23
        Case "PAGEUP": WoVkByName = &H21
        Case "PAGEDOWN": WoVkByName = &H22
        Case "F1": WoVkByName = &H70
        Case "F2": WoVkByName = &H71
        Case "F3": WoVkByName = &H72
        Case "F4": WoVkByName = &H73
        Case "F5": WoVkByName = &H74
        Case "F6": WoVkByName = &H75
        Case "F7": WoVkByName = &H76
        Case "F8": WoVkByName = &H77
        Case "F9": WoVkByName = &H78
        Case "F10": WoVkByName = &H79
        Case "F11": WoVkByName = &H7A
        Case "F12": WoVkByName = &H7B
        Case Else
            WoVkByName = WoNum(s, ok)
    End Select
End Function

Public Function SendWindowMessageJson(ByVal hwnd As Long, ByVal action As String, _
        ByVal message As String, ByVal vk As String, ByVal wparam As String, _
        ByVal lparam As String, ByVal presses As Long, ByVal post As Boolean, _
        ByVal confirm As Boolean) As String
    On Error Resume Next
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("hwnd") = hwnd
    root("action") = action
    root("post") = post
    If presses <= 0 Then presses = 1
    If presses > 20 Then presses = 20
    root("presses") = presses
    If hwnd = 0 Then
        root("ok") = False
        root("reason") = "hwnd_required"
        root("advice") = "pass the hwnd reported by find_gui / read_active_dialogs / ui_text"
        SendWindowMessageJson = JsonStringify(root)
        Exit Function
    End If
    If WoIsWindow(hwnd) = 0 Then
        root("ok") = False
        root("reason") = "window_gone"
        SendWindowMessageJson = JsonStringify(root)
        Exit Function
    End If
    If Not WoMine(hwnd) Then
        root("ok") = False
        root("reason") = "not_this_process"
        root("ownerPid") = WoPidOf(hwnd)
        root("advice") = "this primitive only sends messages to windows of the VB6 process that hosts the add-in"
        SendWindowMessageJson = JsonStringify(root)
        Exit Function
    End If
    root("cls") = WoCls(hwnd)
    root("title") = WoText(hwnd)
    Dim a As String
    a = LCase$(Trim$(action))
    ' TASK-P1-076 (contract v1.6.0): HIGH-RISK message guard. WM_CLOSE / WM_QUIT /
    ' WM_SYSCOMMAND used to go out unguarded, so a message-only primitive could close a
    ' window - or ask the whole thread (the IDE) to quit - by accident. They now require
    ' an EXPLICIT confirm=true. Everything else is unchanged.
    If Len(a) = 0 Or a = "message" Then
        Dim hz As Long, hzOk As Boolean
        hz = -1
        hzOk = False
        If Len(Trim$(message)) > 0 Then
            If IsNumeric(Left$(Trim$(message), 1)) Then
                hz = WoNum(message, hzOk)
            ElseIf LCase$(Left$(Trim$(message), 2)) = "0x" Or LCase$(Left$(Trim$(message), 2)) = "&h" Then
                hz = WoNum(message, hzOk)
            Else
                hz = WoMsgByName(message, hzOk)
            End If
        End If
        If hzOk Then
            If hz = WM_CLOSE Or hz = &H12 Or hz = WM_SYSCOMMAND Then
                If Not confirm Then
                    root("ok") = False
                    root("reason") = "high_risk_message"
                    root("messageAsked") = message
                    root("msg") = hz
                    root("confirmRequired") = True
                    root("risk") = "WM_CLOSE closes that window, WM_QUIT asks the thread to quit (the IDE), WM_SYSCOMMAND can minimize/maximize/close it"
                    root("advice") = "pass confirm=true only when acting on THAT window IS the intent; nothing was sent"
                    SendWindowMessageJson = JsonStringify(root)
                    Exit Function
                End If
                root("highRiskConfirmed") = True
            End If
        End If
    End If
    Dim plan As Object
    Set plan = New Collection
    Dim i As Long
    Dim okA As Boolean
    If a = "click" Then
        Dim e1 As Object
        Set e1 = CreateObject("Scripting.Dictionary")
        e1("name") = "BM_CLICK"
        e1("msg") = BM_CLICK
        e1("wparam") = 0
        e1("lparam") = 0
        plan.Add e1
    ElseIf a = "key" Or a = "send_key" Then
        Dim kv As Long
        kv = WoVkByName(vk, okA)
        If Not okA Then
            root("ok") = False
            root("reason") = "bad_key"
            root("key") = vk
            root("advice") = "use a VK name (ENTER ESC TAB F1..F12 UP DOWN LEFT RIGHT A..Z 0..9) or a number such as 0x0D"
            SendWindowMessageJson = JsonStringify(root)
            Exit Function
        End If
        root("vk") = kv
        For i = 1 To presses
            Dim e2 As Object
            Set e2 = CreateObject("Scripting.Dictionary")
            e2("name") = "WM_KEYDOWN"
            e2("msg") = WM_KEYDOWN
            e2("wparam") = kv
            e2("lparam") = 1
            plan.Add e2
            Set e2 = CreateObject("Scripting.Dictionary")
            e2("name") = "WM_KEYUP"
            e2("msg") = WM_KEYUP
            e2("wparam") = kv
            e2("lparam") = -1073741823
            plan.Add e2
        Next
    Else
        Dim mv As Long, okM As Boolean
        mv = WoMsgByName(message, okM)
        If Not okM Then mv = WoNum(message, okM)
        If Not okM Then
            root("ok") = False
            root("reason") = "bad_message"
            root("message") = message
            root("advice") = "use a message name (click bm_click wm_command wm_keydown wm_keyup wm_char wm_close wm_mousewheel wm_vscroll wm_syscommand bm_setstate cb_showdropdown em_linescroll) or a number such as 0x0111"
            SendWindowMessageJson = JsonStringify(root)
            Exit Function
        End If
        Dim wv As Long, lv As Long
        Dim okW As Boolean, okL As Boolean
        okW = True: okL = True
        wv = 0: lv = 0
        If Len(Trim$(wparam)) > 0 Then wv = WoNum(wparam, okW)
        If Len(Trim$(lparam)) > 0 Then lv = WoNum(lparam, okL)
        If (Not okW) Or (Not okL) Then
            root("ok") = False
            root("reason") = "bad_wparam_or_lparam"
            root("advice") = "wparam/lparam accept decimal or 0x hex"
            SendWindowMessageJson = JsonStringify(root)
            Exit Function
        End If
        root("message") = mv
        root("wparam") = wv
        root("lparam") = lv
        For i = 1 To presses
            Dim e3 As Object
            Set e3 = CreateObject("Scripting.Dictionary")
            e3("name") = message
            e3("msg") = mv
            e3("wparam") = wv
            e3("lparam") = lv
            plan.Add e3
        Next
    End If
    Dim sent As Object
    Set sent = New Collection
    For i = 1 To plan.Count
        Dim p As Object
        Set p = plan.Item(i)
        Dim r As Long
        r = 0
        If post Then
            r = WoPostMessageW(hwnd, CLng(p("msg")), CLng(p("wparam")), CLng(p("lparam")))
        Else
            r = WoSendMessageW(hwnd, CLng(p("msg")), CLng(p("wparam")), CLng(p("lparam")))
        End If
        Dim s As Object
        Set s = CreateObject("Scripting.Dictionary")
        s("name") = p("name")
        s("msg") = p("msg")
        s("wparam") = p("wparam")
        s("lparam") = p("lparam")
        s("sendResult") = r
        sent.Add s
    Next
    root("ok") = True
    Set root("sent") = sent
    root("windowStillExists") = (WoIsWindow(hwnd) <> 0)
    root("note") = "messages only: SendMessage/PostMessage to a window of this process; no memory is written, nothing is injected and no IAT is patched"
    SendWindowMessageJson = JsonStringify(root)
End Function

' =====================================================================
' TASK-P1-067: optional code-pane wheel - OPT-IN, DEFAULT OFF.
'
' WHY SUBCLASSING AND NOT A HOOK. A global keyboard hook is exactly what the
' user wants to get rid of (it is the suspected cause of the machine-wide input
' freeze), and a mouse hook is still a hook, so NONE is installed. Windows
' delivers WM_MOUSEWHEEL to the FOCUSED window and the VB6 code pane (window
' class "VbaWindow", 内部窗口类清单) ignores it; we chain our
' own window procedure in front of the pane's, for panes of our OWN process only.
'
' UNINSTALL. Every hooked hwnd remembers its previous procedure. Uninstall
' restores it ONLY IF the window's current procedure is still OURS, so a
' third-party add-in that subclassed the same window AFTER us keeps a valid
' chain (we deactivate instead of ripping the chain out), and re-enabling never
' stacks because an already-hooked hwnd is skipped.
' =====================================================================

Private Function WoKeyDown(ByVal vk As Long) As Boolean
    WoKeyDown = (WoGetKeyState(vk) < 0)
End Function

Private Function WoAncestorPane(ByVal h As Long) As Long
    Dim p As Long, guard As Long
    p = h
    guard = 0
    Do While p <> 0 And guard < 12
        If LCase$(WoCls(p)) = "vbawindow" Then
            WoAncestorPane = p
            Exit Function
        End If
        p = WoGetParent(p)
        guard = guard + 1
    Loop
End Function

Private Function WoPrevOf(ByVal h As Long) As Long
    Dim i As Long
    For i = 0 To m_whCount - 1
        If m_whHwnd(i) = h Then
            WoPrevOf = m_whPrev(i)
            Exit Function
        End If
    Next
End Function

' The window procedure. It NEVER raises (On Error Resume Next) and always
' forwards anything it does not consume.
Public Function WheelWndProc(ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
    On Error Resume Next
    If msg = WM_MOUSEWHEEL Then
        If m_whEnabled And Not m_whInProc Then
            If WheelHandle(h, wp, lp) Then
                WheelWndProc = 0
                Err.Clear
                Exit Function
            End If
        End If
    End If
    Dim prev As Long
    prev = WoPrevOf(h)
    If prev <> 0 Then
        WheelWndProc = WoCallWindowProcW(prev, h, msg, wp, lp)
    Else
        WheelWndProc = WoDefWindowProcW(h, msg, wp, lp)
    End If
    Err.Clear
End Function

' True = the wheel was consumed (the caller must answer 0).
Private Function WheelHandle(ByVal h As Long, ByVal wp As Long, ByVal lp As Long) As Boolean
    On Error Resume Next
    ' TASK-P1-088: record EVERY observed WM_MOUSEWHEEL with its raw wParam/lParam, so the
    ' receipt can say whether the message ever arrived at our WndProc (msgSeen) and what
    ' it carried (lastDelta = HiWord(wParam), lastWparam/lastLparam raw).
    m_whEvents = m_whEvents + 1
    m_whLastWparam = wp
    m_whLastLparam = lp
    m_whLastHwnd = h
    m_whLastDelta = HiWord(wp)
    Dim pane As Long
    pane = WoAncestorPane(h)
    If pane = 0 Then
        m_whForwarded = m_whForwarded + 1
        m_whLastReason = "no_ancestor_pane"
        Exit Function
    End If
    m_whLastPane = pane
    m_whLastTitle = WoText(pane)
    ' lParam of WM_MOUSEWHEEL is the SCREEN position of the pointer.
    Dim x As Long, y As Long
    x = LoWord(lp)
    y = HiWord(lp)
    If x = 0 And y = 0 Then
        Dim pt As WO_POINT
        If WoGetCursorPos(pt) <> 0 Then
            x = pt.x
            y = pt.y
        End If
    End If
    ' Only when the pointer really is over THIS code pane.
    Dim rc As WO_RECT
    rc.l = 0: rc.t = 0: rc.r = 0: rc.b = 0
    WoGetWindowRect pane, rc
    If x < rc.l Or x > rc.r Or y < rc.t Or y > rc.b Then
        m_whForwarded = m_whForwarded + 1
        m_whLastReason = "pointer_outside_pane"
        Exit Function
    End If
    Dim delta As Long
    delta = HiWord(wp)
    If delta = 0 Then
        m_whForwarded = m_whForwarded + 1
        m_whLastReason = "zero_delta"
        Exit Function
    End If
    Dim notches As Long
    notches = Abs(delta) \ 120
    If notches < 1 Then notches = 1
    Dim up As Boolean
    up = (delta > 0)
    Dim paging As Boolean, fast As Boolean
    paging = WoKeyDown(VK_CONTROL)
    fast = WoKeyDown(VK_SHIFT)
    Dim wantLines As Long, wantPages As Long
    m_whInProc = True
    If paging Then
        wantPages = notches
        If up Then wantPages = -wantPages
        m_whLastScroll = notches
        m_whLastReason = "handled_paging"
    Else
        Dim lines As Long
        lines = m_whLines
        If fast Then lines = lines * 3
        wantLines = notches * lines
        If up Then wantLines = -wantLines
        m_whLastScroll = notches * lines
        If fast Then
            m_whLastReason = "handled_lines_fast"
        Else
            m_whLastReason = "handled_lines"
        End If
    End If
    ' TASK-P1-067 final fix: move the pane through the VBIDE object model.
    ' WHY: WM_VSCROLL to the REAL code pane was measured to do NOTHING 5/5
    ' (SB_LINEDOWN x2 / SB_PAGEDOWN / SB_LINEUP x2 all left CodePane.TopLine at 1),
    ' while setting TopLine directly is observable and reversible 4/4. Only the
    ' VIEWPORT moves here: the selection and the caret are never touched.
    Dim newTop As Long
    newTop = -1
    m_whLastTopBefore = -1
    m_whTopStatus = ""
    m_whTopErr = ""
    If Not m_whSink Is Nothing Then
        m_whLastTopBefore = m_whSink.WheelPaneTopByCaption(m_whLastTitle)
        newTop = m_whSink.WheelScrollPaneByCaption(m_whLastTitle, wantLines, wantPages)
        m_whResolveVia = m_whSink.WheelResolveVia()
        m_whTopStatus = m_whSink.WheelTopStatus()
        m_whTopErr = m_whSink.WheelTopError()
    End If
    m_whInProc = False
    m_whLastPaneCaption = m_whLastTitle
    m_whLastTopAfter = newTop
    If newTop < 0 Then
        ' TASK-P1-132 (D-130-01): "the pane did not resolve" and "the pane resolved but the
        ' IDE refused the read/write" are DIFFERENT failures and must never share a reason.
        m_whForwarded = m_whForwarded + 1
        If m_whTopStatus = "pane_not_writable_in_break" Then
            m_whLastReason = "pane_not_writable_in_break"
            ' The refusal happened INSIDE the subclassed window procedure. Queue the scroll
            ' so the TIMER (outside the WndProc) retries it, and let the receipt carry the
            ' outcome instead of guessing.
            If Not m_whPendOn Then
                m_whPendOn = True
                m_whPendTitle = m_whLastTitle
                m_whPendLines = wantLines
                m_whPendPages = wantPages
                m_whPendTries = 0
                m_whPendResult = ""
            End If
        Else
            m_whLastReason = "pane_not_found"
        End If
        WheelHandle = False
        Err.Clear
        Exit Function
    End If
    m_whHandled = m_whHandled + 1
    WheelHandle = True
    Err.Clear
End Function

' TASK-P1-088, LEGACY (no longer the behaviour oracle as of TASK-P1-067): first
' visible line of the pane's TEXT child. EM_GETFIRSTVISIBLELINE is the RichEdit
' message (0xCE); TASK-P1-083 measured "RichEdit20A" as the code pane's text
' child class. 0 = the child was not found (the field then carries no information).
' This is a DIFFERENT quantity from get_codepane.topLine (VBIDE CodePane.TopLine,
' 1-based): compare MOVEMENT, not raw equality.
Private Function WheelPaneTop(ByVal pane As Long) As Long
    On Error Resume Next
    ' -1 = CHANNEL UNAVAILABLE (no RichEdit20A child under this pane, so there is nothing
    ' that answers EM_GETFIRSTVISIBLELINE) - NOT "line 0". Measured on 2026-10-09: the code
    ' pane in this session exposes NO RichEdit20A child at all (find_gui cls=RichEdit20A ->
    ' 0 hits), so this reads -1 in practice and the usable pane-top oracle stays VBIDE's
    ' CodePane.TopLine (see paneTopLineNow, stamped by CodeTools).
    WheelPaneTop = -1
    Dim child As Long
    child = WoFindWindowExW(pane, 0, StrPtr("RichEdit20A"), 0)
    If child = 0 Then Exit Function
    WheelPaneTop = WoSendMessageW(child, EM_GETFIRSTVISIBLELINE, 0, 0)
    Err.Clear
End Function

' TASK-P1-088: which of the three candidate reasons the CURRENT wheel state points at.
Private Function WheelReason() As String
    If Not m_whEnabled Then
        WheelReason = "switch_off"
    ElseIf m_whCount = 0 Then
        WheelReason = "no_target_window"
    ElseIf m_whEvents = 0 Then
        WheelReason = "armed_no_message_yet"
    ElseIf m_whHandled = 0 Then
        WheelReason = "message_seen_but_not_handled"
    Else
        WheelReason = m_whLastReason
    End If
End Function

' TASK-P1-088: READ-ONLY wheel snapshot, embedded as the "wheel" object of mcp_info and
' ide_state (pure output - no new tool, no description change, so no re-freeze). It exists
' to tell three otherwise indistinguishable failure reasons apart:
'   armed=False               -> (1) nothing was subclassed (switch off / no pane yet)
'   armed=True and msgSeen=0  -> (2) the WM_MOUSEWHEEL never reached our WndProc
'   armed=True and msgSeen>0  -> (3) our handler ran: compare lastDelta / scrollLines with
'                                lastTopLineBefore/After to see whether the pane moved
Public Function WheelStateObj() As Object
    Dim d As Object, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    d("switchOn") = m_whEnabled
    d("installed") = m_whInstalled
    d("armed") = (m_whCount > 0)
    d("subclassCount") = m_whCount
    ' TASK-P1-067: every hooked window IS a code pane now (WheelScan refuses any
    ' VbaWindow whose title is not a live CodePane caption), so these fields can no
    ' longer point at the Immediate window - the misleading observability that sent
    ' an earlier measurement to the wrong window.
    d("targetHwnd") = 0
    d("targetTitle") = ""
    d("targetCaption") = ""
    If m_whCount > 0 Then
        d("targetHwnd") = m_whHwnd(0)
        d("targetTitle") = m_whTitle(0)
        d("targetCaption") = m_whTitle(0)
    End If
    d("skippedCount") = m_whSkipCount
    d("lastPaneCaption") = m_whLastPaneCaption
    Dim mc As Object
    Set mc = New Collection
    If Not m_whCaps Is Nothing Then
        Dim ci As Long
        For ci = 1 To m_whCaps.Count
            mc.Add m_whCaps.Item(ci)
        Next
    End If
    Set d("mappedCaptions") = mc
    Dim sk As Object
    Set sk = New Collection
    Dim si As Long
    For si = 0 To m_whSkipCount - 1
        Dim se As Object
        Set se = CreateObject("Scripting.Dictionary")
        se("hwnd") = m_whSkipHwnd(si)
        se("cls") = m_whSkipCls(si)
        se("title") = m_whSkipTitle(si)
        se("reason") = m_whSkipReason(si)
        sk.Add se
    Next
    Set d("skipped") = sk
    d("msgSeen") = m_whEvents
    d("handled") = m_whHandled
    d("forwarded") = m_whForwarded
    d("scans") = m_whScans
    d("linesPerNotch") = m_whLines
    d("scrollLines") = m_whLastScroll
    d("lastDelta") = m_whLastDelta
    d("lastWparam") = m_whLastWparam
    d("lastLparam") = m_whLastLparam
    d("lastHwnd") = m_whLastHwnd
    d("lastPaneHwnd") = m_whLastPane
    d("lastPaneTitle") = m_whLastTitle
    d("lastTopLineBefore") = m_whLastTopBefore
    d("lastTopLineAfter") = m_whLastTopAfter
    ' TASK-P1-132 (D-130-02): -1 must never look like "line 0" or like "did not move".
    ' TASK-P1-132 (D-130-02): -1 must never look like "line 0" or like "did not move".
    d("lastTopLineBeforeKnown") = (m_whLastTopBefore > 0)
    d("lastTopLineBeforeNote") = ""
    If m_whLastTopBefore < 0 Then
        d("lastTopLineBeforeNote") = "UNAVAILABLE, not line 0. The exact cause is in lastTopStatus/lastTopError: pane_not_found = m_whLastTitle did not resolve to a CodePane; pane_not_writable_in_break = the pane WAS resolved but the IDE refused the TopLine read/write (measured 2026-10-13 in break mode; the real VB error is carried in lastTopError). Read paneTopLineNow (ActiveCodePane.TopLine, stamped by CodeTools) for the CURRENT line instead of comparing this field."
    End If
    ' TASK-P1-132 (D-130-01): why the caption table was empty/partial on the last scan.
    d("captionsSource") = m_whCapsSource
    d("captionsCount") = m_whCapsCount
    d("lastResolveVia") = m_whResolveVia
    d("lastTopStatus") = m_whTopStatus
    d("lastTopError") = m_whTopErr
    d("pendingScroll") = m_whPendOn
    d("pendingScrollResult") = m_whPendResult
    d("topLineSpace") = "TASK-P1-067: lastTopLineBefore/After are VBIDE CodePane.TopLine (1-based) read back around the scroll; -1 means NO VALUE and the reason is in lastTopStatus/lastTopError: pane_not_found = the title resolved to no CodePane (forwarded untouched), pane_not_writable_in_break = the pane WAS resolved but the IDE refused the TopLine read/write (real VB error in lastTopError; the scroll is retried outside the window procedure by the timer - see pendingScroll/pendingScrollResult). The old EM_GETFIRSTVISIBLELINE channel is gone from these two fields because the code pane exposes no RichEdit20A child on this IDE (it read -1 and was never the behaviour oracle). paneTopLineNow is the ActiveCodePane.TopLine stamped by CodeTools."
    d("lastReason") = m_whLastReason
    d("reason") = WheelReason()
    Dim hs As Object
    Set hs = New Collection
    For i = 0 To m_whCount - 1
        Dim he As Object
        Set he = CreateObject("Scripting.Dictionary")
        he("hwnd") = m_whHwnd(i)
        he("cls") = WoCls(m_whHwnd(i))
        he("title") = WoText(m_whHwnd(i))
        he("caption") = m_whTitle(i)
        hs.Add he
    Next
    Set d("hooked") = hs
    Set WheelStateObj = d
End Function

Public Function WheelStateJson() As String
    WheelStateJson = JsonStringify(WheelStateObj())
End Function

Private Function WheelHooked(ByVal h As Long) As Boolean
    Dim i As Long
    For i = 0 To m_whCount - 1
        If m_whHwnd(i) = h Then
            WheelHooked = True
            Exit Function
        End If
    Next
End Function

' TASK-P1-067: h may only be hooked when cap is a live CodePane caption (see
' WheelScan). The caption is stored next to the hwnd because it IS the key the
' scroll needs - and it is read here, in the scan, never inside the WndProc
' (enumeration/COM in a window procedure is a re-entrancy hazard).
Private Sub WheelHook(ByVal h As Long, ByVal cap As String)
    On Error Resume Next
    If h = 0 Then Exit Sub
    If Len(cap) = 0 Then Exit Sub
    If WheelHooked(h) Then Exit Sub
    If m_whCount >= WHEEL_MAX Then Exit Sub
    Dim prev As Long
    ' The second SetWindowLongW returns the procedure that is in place after the
    ' first one - i.e. OURS - which is the exact value the uninstall self-check
    ' needs. (AddressOf is only ever used as an API argument, never in an
    ' expression, which is the portable VB6 form.)
    prev = WoSetWindowLongW(h, GWLP_WNDPROC, AddressOf WheelWndProc)
    m_whSelf = WoSetWindowLongW(h, GWLP_WNDPROC, AddressOf WheelWndProc)
    If prev = 0 Then Exit Sub
    m_whHwnd(m_whCount) = h
    m_whPrev(m_whCount) = prev
    m_whTitle(m_whCount) = cap
    m_whCount = m_whCount + 1
    m_whInstalled = True
End Sub

' TASK-P1-067: is cap the caption of a code pane that is open RIGHT NOW?
Private Function WheelCaptionMapped(ByVal cap As String, ByVal caps As Object) As Boolean
    On Error Resume Next
    If Len(cap) = 0 Then Exit Function
    If caps Is Nothing Then Exit Function
    Dim i As Long
    For i = 1 To caps.Count
        ' verbatim, case-sensitive: the hwnd's GetWindowText is exactly
        ' CodePane.Window.Caption. The "(Code)" suffix is NOT parsed - localised
        ' IDEs change it, the full string is what both sides carry.
        If StrComp(CStr(caps.Item(i)), cap, vbBinaryCompare) = 0 Then
            WheelCaptionMapped = True
            Exit Function
        End If
    Next
End Function

' TASK-P1-067: exactly one VbaWindow is a code pane; everything else with the
' SAME class (the Immediate / Watch / Locals windows) is skipped and recorded,
' so the receipt can never again show a non-code window as the target.
'
' TASK-P1-130 (measured 2026-10-13, pre-fix build D454878E): the VBE caption table can come
' back EMPTY while the project is running/broken - the debugger opens the code pane of the
' module it stopped in, and that pane's window need never have existed during design. In
' break mode the watchdog KEPT SCANNING (scans 2 -> 5) yet WheelPaneCaptions() yielded 0
' captions, so every VbaWindow - including the pane showing the current statement - was
' recorded "not_a_code_pane" and never subclassed; a WM_MOUSEWHEEL posted to that pane never
' reached our WndProc (msgSeen did not move) and the wheel did nothing. With nothing to
' classify against, refusing to hook IS the defect. The scroll path re-resolves the caption
' at message time and FORWARDS anything that does not resolve to a CodePane, so hooking an
' unclassified VbaWindow cannot change that window's own behaviour. With a NON-empty table
' the old strict rule is unchanged (P1-067 criterion 4 evidence still holds).
Private Function WheelTableEmpty(ByVal caps As Object) As Boolean
    WheelTableEmpty = True
    If caps Is Nothing Then Exit Function
    If caps.Count > 0 Then WheelTableEmpty = False
End Function

Private Sub WheelConsider(ByVal h As Long, ByVal caps As Object)
    On Error Resume Next
    Dim cap As String
    cap = WoText(h)
    If WheelCaptionMapped(cap, caps) Then
        ' TASK-P1-130: this add-in HAS positively identified a code pane at least once.
        ' ONLY that arms the empty-table fallback below (see m_whEverVerified).
        m_whEverVerified = True
        WheelHook h, cap
    ElseIf Len(cap) > 0 And m_whEverVerified And WheelTableEmpty(caps) Then
        ' TASK-P1-130: the table is empty AND we already know what a code pane looks like
        ' here -> hook it any way. TASK-P1-132 (D-130-01) settles what happens next, and the
        ' receipt says which case it was: the pane could not be resolved at all ->
        ' lastReason "pane_not_found" (the wheel is forwarded untouched, no pane scrolls);
        ' the pane WAS resolved but the IDE refused the TopLine read/write ->
        ' "pane_not_writable_in_break", carrying the real VB error in lastTopError, and the
        ' scroll is retried OUTSIDE this window procedure by the timer (see
        ' pendingScroll/pendingScrollResult); a successful scroll is "handled_lines".
        ' The Immediate/Watch/Locals windows still behave exactly as before, because the
        ' only fallback that can map a window onto the ACTIVE pane requires the window
        ' title to actually CONTAIN that pane's component name.
        WheelHook h, cap
    Else
        WheelSkip h, cap, "not_a_code_pane"
    End If
    Err.Clear
End Sub

Private Sub WheelSkip(ByVal h As Long, ByVal cap As String, ByVal why As String)
    If m_whSkipCount >= WHEEL_SKIP_MAX Then Exit Sub
    m_whSkipHwnd(m_whSkipCount) = h
    m_whSkipCls(m_whSkipCount) = WoCls(h)
    m_whSkipTitle(m_whSkipCount) = cap
    m_whSkipReason(m_whSkipCount) = why
    m_whSkipCount = m_whSkipCount + 1
End Sub

' Drop entries whose window is gone, then hook every code pane that is not yet
' hooked. Hook the same hwnd twice is impossible (that is the no-stacking rule).
Public Sub WheelScan()
    On Error Resume Next
    Dim i As Long, j As Long
    i = 0
    Do While i < m_whCount
        If WoIsWindow(m_whHwnd(i)) = 0 Then
            For j = i To m_whCount - 2
                m_whHwnd(j) = m_whHwnd(j + 1)
                m_whPrev(j) = m_whPrev(j + 1)
                m_whTitle(j) = m_whTitle(j + 1)
            Next
            m_whTitle(m_whCount - 1) = ""
            m_whCount = m_whCount - 1
        Else
            i = i + 1
        End If
    Loop
    m_whScans = m_whScans + 1
    ' TASK-P1-067: (re)build the code-pane caption table for THIS scan. It is the
    ' ONLY thing that distinguishes a code pane from the Immediate / Watch / Locals
    ' windows, which carry the same window class "VbaWindow". Cost is one late-bound
    ' call per scan, not per message: nothing COM happens in the window procedure.
    Set m_whCaps = Nothing
    m_whSkipCount = 0
    If Not m_whSink Is Nothing Then
        Set m_whCaps = m_whSink.WheelPaneCaptions()
        ' TASK-P1-132 (D-130-01): publish WHY the table looks the way it does. In break
        ' mode the VBE.CodePanes collection yields nothing and CodeTools falls back to the
        ' active pane; that is now visible instead of looking like "no panes exist".
        m_whCapsSource = m_whSink.WheelPaneCaptionsSource()
    End If
    m_whCapsCount = 0
    If Not m_whCaps Is Nothing Then m_whCapsCount = m_whCaps.Count
    Dim tops() As Long, nTops As Long, t As Long
    WoScanTops
    nTops = m_woCount
    If nTops > WO_MAX Then nTops = WO_MAX
    ReDim tops(0 To nTops)
    For t = 0 To nTops - 1
        tops(t) = m_woHwnd(t)
    Next
    For t = 0 To nTops - 1
        Dim h As Long
        h = tops(t)
        If h <> 0 Then
            If LCase$(WoCls(h)) = "vbawindow" Then WheelConsider h, m_whCaps
            WoScanKids h
            Dim k As Long
            For k = 0 To m_woCount - 1
                If LCase$(WoCls(m_woHwnd(k))) = "vbawindow" Then WheelConsider m_woHwnd(k), m_whCaps
            Next
        End If
    Next
    m_whLastScan = WoGetTickCount()
    Err.Clear
End Sub

' Throttled rescan; safe to call from the watchdog timer and the message pump.
Public Sub WheelTick()
    ' TASK-P1-084: the armed Add Form confirm has to be ticked even while the wheel
    ' is OFF (the default), so it runs BEFORE the wheel's own early exit. This is the
    ' watchdog's path (ServerForm.tmrWatch) - and the watchdog's WM_TIMER is dispatched
    ' by the modal loop itself, which is precisely why the click can happen at all.
    WoConfirmTick
    ' TASK-P1-132 (D-130-01): run a scroll the WndProc could not perform, from the TIMER
    ' path (outside the subclassed window procedure), then read the result back.
    WheelPumpPending
    If Not m_whEnabled Then Exit Sub
    If m_whCount > 0 Then
        If WoGetTickCount() - m_whLastScan < 1000 Then Exit Sub
    End If
    WheelScan
End Sub

' TASK-P1-132 (D-130-01): perform a pending pane scroll OUTSIDE the subclassed WndProc.
Public Sub WheelPumpPending()
    On Error Resume Next
    If Not m_whPendOn Then Exit Sub
    If m_whSink Is Nothing Then
        m_whPendOn = False
        m_whPendResult = "abandoned: the VBIDE sink is gone"
        Exit Sub
    End If
    m_whPendTries = m_whPendTries + 1
    Dim nt As Long
    nt = m_whSink.WheelScrollPaneByCaption(m_whPendTitle, m_whPendLines, m_whPendPages)
    Dim st As String
    st = m_whSink.WheelTopStatus()
    If nt >= 0 Then
        m_whPendResult = "scrolled_out_of_wndproc topLine=" & CStr(nt) & " via=" & m_whSink.WheelResolveVia()
        m_whPendOn = False
    ElseIf m_whPendTries >= 3 Then
        m_whPendResult = "still_refused status=" & st & " err=" & m_whSink.WheelTopError()
        m_whPendOn = False
    End If
    Err.Clear
End Sub

' TASK-P1-067: wire the VBIDE side in. WinOps.bas is a STANDARD module with no
' VBIDE reference; CodeTools is the only holder of the VBE, so McpServer hands it
' over here and the two scroll helpers are called late-bound. Called once at
' start-up (ServerHost.Start -> McpServer.Init), before WheelInit.
Public Sub WheelSetSink(ByVal o As Object)
    On Error Resume Next
    Set m_whSink = o
    Err.Clear
End Sub

' OPT-IN. DEFAULT OFF: with VB6MCP_WHEEL unset, nothing is subclassed at all.
Public Sub WheelInit()
    On Error Resume Next
    Dim s As String
    s = LCase$(Trim$(Environ$("VB6MCP_WHEEL")))
    m_whEnabled = (s = "1" Or s = "true" Or s = "on" Or s = "yes")
    If Not m_whEnabled Then
        ' registry twin, same shape as VB6MCP_TRANSPORT / VB6MCP_PORT (TASK-P1-079: advapi32,
        ' because WScript.Shell RegRead/RegWrite is unreliable inside this add-in process)
        s = LCase$(Trim$(WinOps.RegGetString("Software\VB6MCP", "Wheel")))
        If s = "1" Or s = "true" Or s = "on" Or s = "yes" Then m_whEnabled = True
    End If
    m_whLines = 3
    s = Trim$(Environ$("VB6MCP_WHEEL_LINES"))
    If Len(s) = 0 Then
        ' TASK-P1-079: registry twin, same precedence as the Wheel switch itself
        ' (env > HKCU\Software\VB6MCP\WheelLines > 3) - the settings form shows/writes
        ' exactly this chain, so the value it displays is the value that takes effect.
        s = Trim$(McpSettingGet("WheelLines"))
    End If
    If Len(s) > 0 Then
        If IsNumeric(s) Then
            If CLng(Val(s)) >= 1 And CLng(Val(s)) <= 50 Then m_whLines = CLng(Val(s))
        End If
    End If
    If m_whEnabled Then WheelScan
    Err.Clear
End Sub

' CLEAN UNINSTALL. Restore the previous procedure wherever it is still ours;
' where another add-in subclassed after us, leave the chain intact and merely go
' inactive, because replacing the CURRENT procedure there would break THEIR
' chain. Reports what actually happened instead of claiming success.
Public Function WheelShutdownJson() As String
    On Error Resume Next
    m_whEnabled = False
    Dim restored As Long, deferred As Long, i As Long, cur As Long
    restored = 0
    deferred = 0
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim arr As Object
    Set arr = New Collection
    For i = 0 To m_whCount - 1
        If WoIsWindow(m_whHwnd(i)) <> 0 Then
            cur = WoGetWindowLongW(m_whHwnd(i), GWLP_WNDPROC)
            If cur = m_whSelf And m_whSelf <> 0 Then
                WoSetWindowLongW m_whHwnd(i), GWLP_WNDPROC, m_whPrev(i)
                restored = restored + 1
            Else
                deferred = deferred + 1
                Dim e As Object
                Set e = CreateObject("Scripting.Dictionary")
                e("hwnd") = m_whHwnd(i)
                e("reason") = "wndproc_changed_by_another_addin; chain left intact, wheel deactivated"
                arr.Add e
            End If
        End If
    Next
    ' forget every entry: a later re-enable starts from a clean list
    d("wheelEnabled") = False
    d("hadInstalled") = m_whInstalled
    m_whCount = 0
    m_whInstalled = False
    ' TASK-P1-130: a later re-enable must re-earn the empty-table fallback
    m_whEverVerified = False
    ' TASK-P1-067: drop the VBIDE sink too, so nothing can call back into the
    ' add-in after it has been unloaded.
    Set m_whSink = Nothing
    Set m_whCaps = Nothing
    m_whSkipCount = 0
    d("restored") = restored
    d("deferred") = deferred
    Set d("deferredDetail") = arr
    d("note") = "no hook is installed at any time: the wheel is a window subclass on code panes of this process only"
    WheelShutdownJson = JsonStringify(d)
    Err.Clear
End Function
