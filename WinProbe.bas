Attribute VB_Name = "WinProbe"
Option Explicit

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsHungAppWindow Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetClassNameW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function SendMessageTimeoutW Lib "user32" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, ByRef lpdwResult As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef lppe As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

' =====================================================================
' Card TASK-P1-017 (contract v1.1.0 section 4): find_gui
' Enumerate THIS process's top-level windows and all their children, filter by
' keyword / class / visibility, and report geometry plus the parent title chain.
' AddressOf callbacks are only legal in STANDARD modules, which is why this lives
' here and not in CodeTools.cls. Every declare is alias-named so this block can
' never collide with the declarations the rest of this module already uses.
' NOTE: these belong in the DECLARATION SECTION - VB6 rejects a Declare that
' appears after any procedure ("only comments may follow End Function").
' =====================================================================
Private Type FG_RECT
    l As Long
    t As Long
    r As Long
    b As Long
End Type

Private Const FG_MAX As Long = 800
Private m_fgHwnd(0 To FG_MAX - 1) As Long
Private m_fgCount As Long

Private Declare Function FgEnumWindows Lib "user32" Alias "EnumWindows" (ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function FgEnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal h As Long, ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function FgGetWindowTextW Lib "user32" Alias "GetWindowTextW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function FgGetClassNameW Lib "user32" Alias "GetClassNameW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function FgIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal h As Long) As Long
Private Declare Function FgIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal h As Long) As Long
Private Declare Function FgGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal h As Long, ByRef rc As FG_RECT) As Long
Private Declare Function FgGetParent Lib "user32" Alias "GetParent" (ByVal h As Long) As Long
Private Declare Function FgGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal h As Long, ByRef pid As Long) As Long
Private Declare Function FgGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
' =====================================================================

' ---- contract v1.1.0 section 5 (TASK-P1-019): ui_text ----
' Three channels, tried in order: UI Automation -> MSAA (IAccessible through
' AccessibleObjectFromWindow) -> WM_GETTEXT. A NEGATIVE result is a legitimate
' delivery: every channel reports its raw return code / error, so "no accessible
' text" can be reproduced instead of trusted.
' Declares belong in the DECLARATION SECTION (VB6 rejects one that follows a
' procedure), and every Win32 entry point is alias-named to avoid collisions.
Private Type UIT_GUID
    D1 As Long
    D2 As Integer
    D3 As Integer
    D4 As Byte
    D5 As Byte
    D6 As Byte
    D7 As Byte
    D8 As Byte
    D9 As Byte
    D10 As Byte
    D11 As Byte
End Type

Private Declare Function UiSendMessageW Lib "user32" Alias "SendMessageW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function UiGetWindow Lib "user32" Alias "GetWindow" (ByVal h As Long, ByVal cmd As Long) As Long
' NOTE (measured twice on the real IDE): neither "ByRef ppvObject As Object" nor
' "ppvObject As Any" lets oleacc fill a VB6 object variable - the API answers S_OK
' while the variable stays Nothing, even for a standard Button that certainly has an
' IAccessible. So the channel asks for a raw pointer and copies it into the object
' slot itself; that is the only form that worked.
Private Declare Function UiAccessibleObjectFromWindow Lib "oleacc" (ByVal hwnd As Long, ByVal dwId As Long, riid As Any, ByRef ppvObject As Long) As Long
Private Declare Sub UiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal n As Long)
Private Declare Function UiIIDFromString Lib "ole32" (ByVal lpsz As Long, ByRef lpiid As UIT_GUID) As Long

Private Const UiGW_CHILD As Long = 5
Private Const UiGW_HWNDNEXT As Long = 2
Private Const UiWM_GETTEXT As Long = 13
Private Const UiWM_GETTEXTLENGTH As Long = 14
Private Const UiOBJID_CLIENT As Long = -4

Private Const WM_NULL As Long = 0
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Const TH32CS_SNAPPROCESS As Long = &H2

Private m_pid As Long
Private m_count As Long
Private m_h(0 To 255) As Long
Private m_title(0 To 255) As String
Private m_class(0 To 255) As String
Private m_vis(0 To 255) As Long

Public Function CurrentProcessId() As Long
    CurrentProcessId = GetCurrentProcessId()
End Function

Private Function ExeNameOf(ByRef pe As PROCESSENTRY32) As String
    Dim s As String, z As Long
    s = pe.szExeFile
    z = InStr(s, vbNullChar)
    If z > 0 Then ExeNameOf = Left$(s, z - 1) Else ExeNameOf = s
End Function

Public Function FindProcessIdByName(ByVal exeName As String) As Long
    Dim snap As Long, pe As PROCESSENTRY32, ok As Long
    FindProcessIdByName = 0
    snap = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    If snap = -1 Or snap = 0 Then Exit Function
    pe.dwSize = Len(pe)
    ok = Process32First(snap, pe)
    Do While ok <> 0
        If StrComp(ExeNameOf(pe), exeName, vbTextCompare) = 0 Then
            FindProcessIdByName = pe.th32ProcessID
            CloseHandle snap
            Exit Function
        End If
        ok = Process32Next(snap, pe)
    Loop
    CloseHandle snap
End Function

Private Function WinText(ByVal h As Long) As String
    Dim b As String, z As Long
    b = String$(512, vbNullChar)
    GetWindowTextW h, StrPtr(b), 512
    z = InStr(b, vbNullChar)
    If z > 0 Then WinText = Left$(b, z - 1) Else WinText = ""
End Function

Private Function WinClass(ByVal h As Long) As String
    Dim b As String, z As Long
    b = String$(256, vbNullChar)
    GetClassNameW h, StrPtr(b), 256
    z = InStr(b, vbNullChar)
    If z > 0 Then WinClass = Left$(b, z - 1) Else WinClass = ""
End Function

Private Function EnumProc(ByVal hWnd As Long, ByVal lParam As Long) As Long
    Dim p As Long
    GetWindowThreadProcessId hWnd, p
    If p = m_pid Then
        If m_count < 256 Then
            m_h(m_count) = hWnd
            m_title(m_count) = WinText(hWnd)
            m_class(m_count) = WinClass(hWnd)
            m_vis(m_count) = IsWindowVisible(hWnd)
            m_count = m_count + 1
        End If
    End If
    EnumProc = 1
End Function

Private Sub ScanWindows(ByVal pid As Long)
    m_pid = pid
    m_count = 0
    If pid <> 0 Then EnumWindows AddressOf EnumProc, 0
End Sub

Public Function CountTopWindows(ByVal pid As Long) As Long
    ScanWindows pid
    CountTopWindows = m_count
End Function

Public Function CountDialogs(ByVal pid As Long) As Long
    Dim i As Long, n As Long
    ScanWindows pid
    For i = 0 To m_count - 1
        If m_class(i) = "#32770" Then n = n + 1
    Next
    CountDialogs = n
End Function

Public Function MainWindowOf(ByVal pid As Long) As Long
    Dim i As Long, best As Long, bestLen As Long
    ScanWindows pid
    best = 0
    bestLen = 0
    For i = 0 To m_count - 1
        If m_vis(i) <> 0 Then
            If Len(m_title(i)) > bestLen Then
                bestLen = Len(m_title(i))
                best = m_h(i)
            End If
        End If
    Next
    MainWindowOf = best
End Function

Public Function IsResponsive(ByVal hWnd As Long, ByVal timeoutMs As Long) As Boolean
    Dim r As Long
    If hWnd = 0 Then Exit Function
    If timeoutMs <= 0 Then timeoutMs = 2000
    If SendMessageTimeoutW(hWnd, WM_NULL, 0, 0, SMTO_ABORTIFHUNG, timeoutMs, r) <> 0 Then IsResponsive = True
End Function

Public Function IsHung(ByVal hWnd As Long) As Boolean
    If hWnd = 0 Then Exit Function
    If IsHungAppWindow(hWnd) <> 0 Then IsHung = True
End Function

Private Function JsonBool(ByVal b As Boolean) As String
    If b Then JsonBool = "true" Else JsonBool = "false"
End Function

Public Function WindowsJson(ByVal pid As Long, ByVal maxN As Long) As String
    Dim i As Long, s As String, n As Long
    ScanWindows pid
    If maxN <= 0 Or maxN > m_count Then maxN = m_count
    s = "["
    For i = 0 To maxN - 1
        If i > 0 Then s = s & ","
        s = s & "{" & Chr$(34) & "class" & Chr$(34) & ":" & Jq(m_class(i)) & "," & Chr$(34) & "title" & Chr$(34) & ":" & Jq(m_title(i)) & "," & Chr$(34) & "visible" & Chr$(34) & ":" & JsonBool(m_vis(i) <> 0) & "}"
    Next
    s = s & "]"
    WindowsJson = s
End Function

Public Function ExecutionStateJson(ByVal pid As Long) As String
    Dim h As Long, resp As Boolean, hg As Boolean
    h = MainWindowOf(pid)
    resp = False
    hg = False
    If h > 0 Then
        resp = IsResponsive(h, 2000)
        hg = IsHung(h)
    End If
    ExecutionStateJson = "{" & Chr$(34) & "pid" & Chr$(34) & ":" & CStr(pid) & "," & Chr$(34) & "exists" & Chr$(34) & ":" & JsonBool(pid > 0) & "," & Chr$(34) & "mainWindow" & Chr$(34) & ":" & CStr(h) & "," & Chr$(34) & "responsive" & Chr$(34) & ":" & JsonBool(resp) & "," & Chr$(34) & "hung" & Chr$(34) & ":" & JsonBool(hg) & "," & Chr$(34) & "topWindows" & Chr$(34) & ":" & CStr(CountTopWindows(pid)) & "," & Chr$(34) & "dialogs" & Chr$(34) & ":" & CStr(CountDialogs(pid)) & "}"
End Function

Private Function Jq(ByVal v As String) As String
    ' TASK-P1-121 item 4 (P4-095): delegate to the shared JSON string quoter.
    ' The old two-Replace body escaped ONLY the backslash and the double quote, so a
    ' window title/text containing LF, CR, TAB or 0x01 produced INVALID JSON and the
    ' whole find_gui answer became unparseable. JsonLib.Quote escapes control characters
    ' too (it was made Public for this). Declared output change: a string that contains a
    ' control character is now correctly escaped; quotes, backslashes, Chinese text,
    ' 0x7F and the empty string stay byte-identical.
    Jq = JsonLib.Quote(v)
End Function

' ---- Card TASK-P1-017 (contract v1.1.0 section 4): find_gui procedures ----

Private Function FgCollect(ByVal h As Long, ByVal lp As Long) As Long
    If m_fgCount < FG_MAX Then
        m_fgHwnd(m_fgCount) = h
        m_fgCount = m_fgCount + 1
    End If
    FgCollect = 1
End Function

' TASK-P1-120 (D-1118-01): snapshot the desktop's top-level windows into a LOCAL array and
' return the count.  Both callers below used to read the top-level hwnds straight out of m_fgHwnd
' while FgCollect was overwriting that SAME module array from index 0 (m_fgCount had just been
' reset), so the child pass walked a progressively corrupted window sequence: which windows were
' enumerated depended on the desktop Z-order at that instant, the same hwnd could be collected
' many times, and the process's own children were often never reached at all (measured
' TASK-P1-118: 132 items / 93 unique hwnds in one condition, and 26 items covering 0 of the 31
' windows vba6_probe{wnd_probe} could see in another).  Snapshotting first removes the aliasing.
Private Function FgSnapshotTops(ByRef topsArr() As Long) As Long
    Dim n As Long, i As Long
    m_fgCount = 0
    FgEnumWindows AddressOf FgCollect, 0&
    n = m_fgCount
    If n > 0 Then
        ReDim topsArr(0 To n - 1)
        For i = 0 To n - 1
            topsArr(i) = m_fgHwnd(i)
        Next
    End If
    FgSnapshotTops = n
End Function

Private Function FgTextOf(ByVal h As Long) As String
    Dim buf As String, n As Long
    buf = String$(512, vbNullChar)
    n = FgGetWindowTextW(h, StrPtr(buf), 512)
    If n > 0 Then FgTextOf = Left$(buf, n) Else FgTextOf = ""
End Function

Private Function FgClassOf(ByVal h As Long) As String
    Dim buf As String, n As Long
    buf = String$(256, vbNullChar)
    n = FgGetClassNameW(h, StrPtr(buf), 256)
    If n > 0 Then FgClassOf = Left$(buf, n) Else FgClassOf = ""
End Function

' Parent chain of window TITLES, outermost first (bounded: a deep tree must not hang).
Private Function FgPathOf(ByVal h As Long) As String
    Dim s As String, p As Long, guard As Long
    s = ""
    p = FgGetParent(h)
    guard = 0
    Do While p <> 0 And guard < 8
        If Len(FgTextOf(p)) > 0 Then
            If Len(s) > 0 Then s = FgTextOf(p) & " > " & s Else s = FgTextOf(p)
        End If
        p = FgGetParent(p)
        guard = guard + 1
    Loop
    FgPathOf = s
End Function

' The caption of this process's main window (works even while the VBIDE object
' model is suspended by RUN mode).
Public Function MainWindowTitleOf(ByVal pid As Long) As String
    Dim h As Long
    h = MainWindowOf(pid)
    If h = 0 Then
        MainWindowTitleOf = ""
    Else
        MainWindowTitleOf = FgTextOf(h)
    End If
End Function

' showing: -1 = no visibility filter, 0 = hidden only, 1 = visible only.
Public Function FindGuiJson(ByVal keyword As String, ByVal cls As String, ByVal showing As Long) As String
    On Error Resume Next
    Dim k As String, c As String, pid As Long
    k = LCase$(Trim$(keyword))
    c = LCase$(Trim$(cls))
    pid = FgGetCurrentProcessId()
    Dim tops As Long, i As Long, total As Long
    Dim topsArr() As Long
    Dim all() As Long
    ReDim all(0 To FG_MAX * 3 - 1)
    ' TASK-P1-120 (D-1118-01): snapshot the tops BEFORE the child pass reuses m_fgHwnd as
    ' FgCollect's destination - see FgSnapshotTops for the measured failure mechanism.
    tops = FgSnapshotTops(topsArr)
    total = 0
    For i = 0 To tops - 1
        If total <= UBound(all) Then
            all(total) = topsArr(i)
            total = total + 1
        End If
    Next
    m_fgCount = 0
    For i = 0 To tops - 1
        FgEnumChildWindows topsArr(i), AddressOf FgCollect, 0&
    Next
    For i = 0 To m_fgCount - 1
        If total <= UBound(all) Then
            all(total) = m_fgHwnd(i)
            total = total + 1
        End If
    Next
    Dim out As String, n As Long
    out = "["
    n = 0
    Dim dup As Boolean, j As Long
    For i = 0 To total - 1
        Dim h As Long, wpid As Long
        h = all(i)
        wpid = 0
        ' TASK-P1-120 (D-1118-01 (1)): emit each hwnd AT MOST ONCE.  With the aliasing gone the two
        ' passes are already disjoint, but the receipt must be unique by construction rather than by
        ' luck (the aliased version returned 132 items for 93 distinct hwnds: +39 duplicates).
        dup = False
        For j = 0 To i - 1
            If all(j) = h Then
                dup = True
                Exit For
            End If
        Next
        If h <> 0 Then FgGetWindowThreadProcessId h, wpid
        If wpid = pid And h <> 0 And Not dup Then
            Dim tx As String, cn As String, vis As Boolean, en As Boolean
            tx = FgTextOf(h)
            cn = FgClassOf(h)
            vis = (FgIsWindowVisible(h) <> 0)
            en = (FgIsWindowEnabled(h) <> 0)
            Dim hit As Boolean
            hit = True
            If Len(k) > 0 Then
                If InStr(1, LCase$(tx), k, vbTextCompare) = 0 And InStr(1, LCase$(cn), k, vbTextCompare) = 0 Then hit = False
            End If
            If Len(c) > 0 Then
                If LCase$(cn) <> c Then hit = False
            End If
            If showing = 1 And Not vis Then hit = False
            If showing = 0 And vis Then hit = False
            If hit Then
                Dim rc As FG_RECT
                rc.l = 0: rc.t = 0: rc.r = 0: rc.b = 0
                FgGetWindowRect h, rc
                If n > 0 Then out = out & ","
                out = out & "{" & Jq("hwnd") & ":" & CStr(h) & "," & Jq("cls") & ":" & Jq(cn) & "," & Jq("text") & ":" & Jq(tx) & "," & Jq("showing") & ":" & JsonBool(vis) & "," & Jq("enabled") & ":" & JsonBool(en) & "," & Jq("rect") & ":{" & Jq("l") & ":" & CStr(rc.l) & "," & Jq("t") & ":" & CStr(rc.t) & "," & Jq("w") & ":" & CStr(rc.r - rc.l) & "," & Jq("h") & ":" & CStr(rc.b - rc.t) & "}," & Jq("path") & ":" & Jq(FgPathOf(h)) & "}"
                n = n + 1
            End If
        End If
    Next
    FindGuiJson = out & "]"
End Function


' ---- contract v1.1.0 section 5 (TASK-P1-019): ui_text procedures ----

' First window of THIS process whose class matches (top-level first, then any
' descendant). Returns 0 when nothing matches.
Private Function UiFindByClass(ByVal cls As String) As Long
    On Error Resume Next
    UiFindByClass = 0
    Dim want As String
    want = LCase$(Trim$(cls))
    If Len(want) = 0 Then Exit Function
    Dim i As Long, tops As Long
    Dim topsArr() As Long
    ' TASK-P1-120 (D-1118-01): identical aliasing pattern to FindGuiJson - fixed the same way.
    tops = FgSnapshotTops(topsArr)
    For i = 0 To tops - 1
        If LCase$(FgClassOf(topsArr(i))) = want Then
            UiFindByClass = topsArr(i)
            Exit Function
        End If
    Next
    m_fgCount = 0
    For i = 0 To tops - 1
        FgEnumChildWindows topsArr(i), AddressOf FgCollect, 0&
    Next
    For i = 0 To m_fgCount - 1
        If LCase$(FgClassOf(m_fgHwnd(i))) = want Then
            UiFindByClass = m_fgHwnd(i)
            Exit Function
        End If
    Next
End Function

' Depth-limited child-window list: root first, then children breadth-first down to
' 'depth' levels (bounded to 400 nodes so a deep tree cannot hang the server).
Private Sub UiCollectTree(ByVal root As Long, ByVal depth As Long, ByRef hs() As Long, ByRef hn As Long)
    On Error Resume Next
    hn = 0
    ReDim hs(0 To 399)
    If root = 0 Then Exit Sub
    hs(hn) = root
    hn = hn + 1
    Dim levelStart As Long, levelEnd As Long, lvl As Long, i As Long, c As Long
    levelStart = 0
    levelEnd = 1
    For lvl = 0 To depth - 1
        Dim nextStart As Long
        nextStart = hn
        For i = levelStart To levelEnd - 1
            c = UiGetWindow(hs(i), UiGW_CHILD)
            Do While c <> 0
                If hn < 400 Then
                    hs(hn) = c
                    hn = hn + 1
                End If
                c = UiGetWindow(c, UiGW_HWNDNEXT)
            Loop
        Next
        levelStart = nextStart
        levelEnd = hn
        If levelStart >= levelEnd Then Exit For
    Next
End Sub

Private Function UiWmGetText(ByVal h As Long) As String
    On Error Resume Next
    Dim n As Long, buf As String, r As Long
    n = UiSendMessageW(h, UiWM_GETTEXTLENGTH, 0, 0)
    If n <= 0 Then Exit Function
    If n > 4096 Then n = 4096
    buf = String$(n + 1, vbNullChar)
    r = UiSendMessageW(h, UiWM_GETTEXT, n + 1, StrPtr(buf))
    If r > 0 Then UiWmGetText = Left$(buf, r) Else UiWmGetText = ""
End Function

' The ui_text tool body (returns JSON text). Channels: UIA -> MSAA -> WM_GETTEXT.
Public Function UiTextJson(ByVal hwnd As Long, ByVal cls As String, ByVal depth As Long) As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim chan As Object
    Set chan = CreateObject("Scripting.Dictionary")
    If depth <= 0 Then depth = 3
    If depth > 5 Then depth = 5
    Dim target As Long
    Dim resolveHow As String
    target = hwnd
    resolveHow = "hwnd argument"
    If target = 0 Then
        target = UiFindByClass(cls)
        resolveHow = "resolved by cls"
    End If
    d("hwnd") = target
    d("resolve") = resolveHow
    d("depth") = depth
    If target = 0 Then
        d("ok") = False
        d("reason") = "no_accessible_text"
        d("evidence") = "no window of this process matched hwnd=" & CStr(hwnd) & " cls=" & cls
        Set d("channels") = chan
        UiTextJson = JsonStringify(d)
        Exit Function
    End If
    d("cls") = FgClassOf(target)
    d("title") = FgTextOf(target)

    ' ---- channel 1: UI Automation ----
    Dim uiaDetail As String, uiaText As String, uiaObj As Object, uiaEl As Object
    uiaDetail = ""
    uiaText = ""
    Err.Clear
    Set uiaObj = CreateObject("UIAutomationClient.CUIAutomation")
    If Err.Number <> 0 Then
        uiaDetail = "CreateObject(UIAutomationClient.CUIAutomation) failed: " & Err.Description
        Err.Clear
    ElseIf uiaObj Is Nothing Then
        uiaDetail = "CreateObject returned Nothing"
    Else
        Set uiaEl = uiaObj.ElementFromHandle(target)
        If Err.Number <> 0 Then
            uiaDetail = "ElementFromHandle failed (IUIAutomation is not IDispatch, so VB6 late binding cannot call it): " & Err.Description
            Err.Clear
        ElseIf uiaEl Is Nothing Then
            uiaDetail = "ElementFromHandle returned Nothing"
        Else
            uiaText = CStr(uiaEl.CurrentName)
            If Err.Number <> 0 Then
                uiaDetail = "CurrentName failed: " & Err.Description
                Err.Clear
            End If
        End If
    End If
    chan("uia") = uiaDetail
    chan("uiaText") = uiaText

    ' ---- channel 2: MSAA (IAccessible through AccessibleObjectFromWindow) ----
    Dim iid As UIT_GUID
    Dim iidHr As Long, msaaHr As Long
    iidHr = UiIIDFromString(StrPtr("{618736E0-3C3D-11CF-810C-00AA00389B71}"), iid)
    chan("iidHr") = iidHr
    Dim acc As Object
    Dim pAcc As Long
    Set acc = Nothing
    pAcc = 0
    Err.Clear
    msaaHr = UiAccessibleObjectFromWindow(target, UiOBJID_CLIENT, iid, pAcc)
    chan("msaaHr") = msaaHr
    If msaaHr = 0 And pAcc <> 0 Then UiCopyMemory acc, pAcc, 4
    Dim msaaDetail As String, msaaText As String
    msaaDetail = ""
    msaaText = ""
    If msaaHr <> 0 Then
        msaaDetail = "AccessibleObjectFromWindow returned 0x" & Hex$(msaaHr) & " (this window class does not expose MSAA)"
    ElseIf acc Is Nothing Then
        msaaDetail = "AccessibleObjectFromWindow returned S_OK with a NULL pointer: the window answered WM_GETOBJECT with 0, i.e. it exposes no MSAA object"
    Else
        Dim cc As Long, nm As Variant, vv As Variant, rl As Variant
        cc = -1
        Err.Clear
        cc = CLng(acc.accChildCount)
        If Err.Number <> 0 Then
            msaaDetail = "accChildCount failed: " & Err.Description
            Err.Clear
        End If
        chan("accChildCount") = cc
        Err.Clear
        nm = acc.accName(0)
        If Err.Number = 0 Then msaaText = CStr(nm) Else Err.Clear
        chan("accName0") = msaaText
        Err.Clear
        vv = acc.accValue(0)
        If Err.Number = 0 Then chan("accValue0") = CStr(vv) Else Err.Clear
        Err.Clear
        rl = acc.accRole(0)
        If Err.Number = 0 Then chan("accRole0") = CStr(rl) Else Err.Clear
        Dim i As Long, k As Long, cnm As Variant, lim As Long
        k = 0
        If cc > 0 Then
            lim = cc
            If lim > 50 Then lim = 50
            For i = 1 To lim
                Err.Clear
                cnm = acc.accName(i)
                If Err.Number = 0 Then
                    If Len(CStr(cnm)) > 0 Then
                        msaaText = msaaText & vbCrLf & CStr(cnm)
                        k = k + 1
                    End If
                Else
                    Err.Clear
                End If
            Next
        End If
        chan("msaaChildNames") = k
        msaaDetail = "IAccessible read directly (depth 1); child count " & CStr(cc) & ", non-empty child names " & CStr(k)
    End If
    chan("msaa") = msaaDetail

    ' ---- channel 3: WM_GETTEXT over the depth-limited subtree ----
    Dim hs() As Long, hn As Long
    UiCollectTree target, depth, hs, hn
    Dim j As Long, t As String, hitN As Long, winTexts As String
    hitN = 0
    winTexts = ""
    For j = 0 To hn - 1
        t = UiWmGetText(hs(j))
        If Len(Trim$(t)) > 0 Then
            hitN = hitN + 1
            winTexts = winTexts & vbCrLf & "[" & FgClassOf(hs(j)) & " " & CStr(hs(j)) & "] " & t
        End If
    Next
    chan("wmGetTextWindows") = hn
    chan("wmGetTextNonEmpty") = hitN

    Dim src As String, text As String
    src = ""
    text = ""
    If Len(Trim$(uiaText)) > 0 Then
        src = "uia"
        text = uiaText
    ElseIf Len(Trim$(msaaText)) > 0 Then
        src = "msaa"
        text = msaaText
    ElseIf Len(Trim$(winTexts)) > 0 Then
        src = "wm_gettext"
        text = winTexts
    End If
    If Len(src) > 0 Then
        d("ok") = True
        d("source") = src
        d("text") = text
        If src = "wm_gettext" Then
            If hitN = 1 And Len(Trim$(UiWmGetText(target))) > 0 Then
                d("textScope") = "root caption only"
                d("note") = "the only WM_GETTEXT hit was the window's own caption; no descendant text was readable, which is the shape of a self-drawn panel"
            Else
                d("textScope") = "subtree"
            End If
        Else
            d("textScope") = "root accessible object"
        End If
    Else
        d("ok") = False
        d("reason") = "no_accessible_text"
        d("evidence") = "hwnd=" & CStr(target) & " cls=" & CStr(d("cls")) & " depth=" & CStr(depth) & _
            "; uia=" & uiaDetail & _
            "; msaaHr=0x" & Hex$(msaaHr) & " (" & msaaDetail & ")" & _
            "; wm_gettext scanned " & CStr(hn) & " window(s), 0 returned non-empty text"
    End If
    Set d("channels") = chan
    UiTextJson = JsonStringify(d)
End Function
