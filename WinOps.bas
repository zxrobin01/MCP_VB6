Attribute VB_Name = "WinOps"
Option Explicit

' =====================================================================
' (contract v1.5.0): message-only window primitives.
'   read_active_dialogs   - enumerate THIS process's active dialogs + buttons
'   close_active_dialogs  - click a SAFE button of an active dialog
'   send_window_message   - SendMessage/PostMessage to a hwnd of THIS process
' (opt-in, DEFAULT OFF): code-pane mouse wheel by SUBCLASSING the
'   code-pane window (class "VbaWindow"). NO keyboard hook, NO global hook,
'   NO WH_MOUSE at all: the whole feature is SetWindowLong/CallWindowProc on
'   windows of our own process.
'
' MEMORY DISCIPLINE (auditable by grep - see acceptance (5)): the
' ONLY Win32 entry points in this module are the message/enumeration/geometry
' family, the subclassing trio and the advapi32 REGISTRY family
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

' GetGUIThreadInfo: the FOCUSED window of a thread's own queue, independent of which
' thread owns the foreground (GetFocus answers 0 for a background thread).
Private Type WO_GUITHREADINFO
    cbSize As Long
    flags As Long
    hwndActive As Long
    hwndFocus As Long
    hwndCapture As Long
    hwndMenuOwner As Long
    hwndMoveSize As Long
    hwndCaret As Long
    rcCaret As WO_RECT
End Type

' ---- message / enumeration / geometry -------------------------------------
Private Declare Function WoSendMessageW Lib "user32" Alias "SendMessageW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoPostMessageW Lib "user32" Alias "PostMessageW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoEnumWindows Lib "user32" Alias "EnumWindows" (ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function WoEnumChildWindows Lib "user32" Alias "EnumChildWindows" (ByVal h As Long, ByVal cb As Long, ByVal lp As Long) As Long
Private Declare Function WoGetWindowTextW Lib "user32" Alias "GetWindowTextW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
' FIX-11: several VB6 IDE dialogs (owner-draw buttons) answer WM_GETTEXT with ANSI bytes.
' Reading those as UTF-16 produced mojibake like 0x6D60. Read ANSI and decode ONCE (CP_ACP).
Private Declare Function WoGetWindowTextA Lib "user32" Alias "GetWindowTextA" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function WoMultiByteToWideChar Lib "kernel32" Alias "MultiByteToWideChar" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cbMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Declare Function WoGetFocus Lib "user32" Alias "GetFocus" () As Long
' GetGUIThreadInfo is the basis used instead of GetFocus (see WoGetThreadFocus).
Private Declare Function WoGetGUIThreadInfo Lib "user32" Alias "GetGUIThreadInfo" (ByVal idThread As Long, ByRef lpgti As WO_GUITHREADINFO) As Long
Private Const WO_GTI_SIZE As Long = 48

' ---- diagnostic profiler (opt-in: VB6MCP_PERF_DIAG=1) --------------------
' QPC is read as TWO Longs: a Currency argument would be scaled by 10000 and a
' LARGE_INTEGER is not expressible in VB6.
Private Type WO_QPC
    lo As Long
    hi As Long
End Type
Private Declare Function WoQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef lpPerformanceCount As WO_QPC) As Long
Private Declare Function WoQueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef lpFrequency As WO_QPC) As Long
Private Const WO_PF_SEGS As Long = 33
Private Const WO_PF_WINDOW_US As Double = 5000000#

' Per-window THREAD CPU time: QPC wall time cannot separate a blocking wait from work,
' and the idle-CPU question needs exactly that difference. GetThreadTimes = 100-ns units.
Private Type WO_FILETIME
    lo As Long
    hi As Long
End Type
Private Declare Function WoGetCurrentThread Lib "kernel32" Alias "GetCurrentThread" () As Long
Private Declare Function WoGetThreadTimes Lib "kernel32" Alias "GetThreadTimes" (ByVal hThread As Long, ByRef lpCreationTime As WO_FILETIME, ByRef lpExitTime As WO_FILETIME, ByRef lpKernelTime As WO_FILETIME, ByRef lpUserTime As WO_FILETIME) As Long

' environment{} (mcp_info only; NEVER on a timer).
Private Declare Function WoGetACP Lib "kernel32" Alias "GetACP" () As Long
Private Declare Function WoGetUserDefaultUILanguage Lib "kernel32" Alias "GetUserDefaultUILanguage" () As Long
Private Declare Function WoGetModuleFileNameW Lib "kernel32" Alias "GetModuleFileNameW" (ByVal hModule As Long, ByVal lpFilename As Long, ByVal nSize As Long) As Long
Private Declare Function WoGetFileVersionInfoSizeW Lib "version" Alias "GetFileVersionInfoSizeW" (ByVal lptstrFilename As Long, ByRef lpdwHandle As Long) As Long
Private Declare Function WoGetFileVersionInfoW Lib "version" Alias "GetFileVersionInfoW" (ByVal lptstrFilename As Long, ByVal dwHandle As Long, ByVal dwLen As Long, ByVal lpData As Long) As Long
Private Declare Function WoVerQueryValueW Lib "version" Alias "VerQueryValueW" (ByVal pBlock As Long, ByVal lpSubBlock As Long, ByRef lplpBuffer As Long, ByRef puLen As Long) As Long
Private Declare Function WoLstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function WoLStrCpyW Lib "kernel32" Alias "lstrCpyW" (ByVal lpString1 As Long, ByVal lpString2 As Long) As Long

' BCP-47 shape of the UI language ("zh-CN"). GetUserDefaultUILanguage gives a LANGID; the
' suite asserts a language TAG, so the LANGID is resolved through GetLocaleInfoW.
Private Declare Function WoGetLocaleInfoW Lib "kernel32" Alias "GetLocaleInfoW" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As Long, ByVal cchData As Long) As Long
' LOCALE_SNAME (Vista+) hands back the COMPLETE tag, e.g. "zh-CN" / "zh-Hans-CN".
Private Const WO_LOCALE_SNAME As Long = &H5C
Private Declare Function WoGetClassNameW Lib "user32" Alias "GetClassNameW" (ByVal h As Long, ByVal buf As Long, ByVal n As Long) As Long
Private Declare Function WoGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal h As Long, ByRef rc As WO_RECT) As Long
Private Declare Function WoGetParent Lib "user32" Alias "GetParent" (ByVal h As Long) As Long
Private Declare Function WoGetWindow Lib "user32" Alias "GetWindow" (ByVal h As Long, ByVal cmd As Long) As Long
Private Declare Function WoIsWindow Lib "user32" Alias "IsWindow" (ByVal h As Long) As Long
Private Declare Function WoIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal h As Long) As Long
Private Declare Function WoIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal h As Long) As Long
Private Declare Function WoGetDlgCtrlID Lib "user32" Alias "GetDlgCtrlID" (ByVal h As Long) As Long
Private Declare Function WoSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hWnd As Long) As Long
Private Declare Function WoGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
Private Declare Function WoGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal h As Long, ByRef pid As Long) As Long
Private Declare Function WoSendMessageTimeoutW Lib "user32" Alias "SendMessageTimeoutW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long, ByVal flags As Long, ByVal timeout As Long, ByRef result As Long) As Long
Private Declare Function WoIsHungAppWindow Lib "user32" Alias "IsHungAppWindow" (ByVal h As Long) As Long
Private Const SMTO_BLOCK As Long = &H1
Private Const SMTO_ABORTIFHUNG As Long = &H2
Private Declare Function WoGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function WoGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function WoGetGuiResources Lib "user32" Alias "GetGuiResources" (ByVal hProcess As Long, ByVal uiFlags As Long) As Long
Private Declare Function WoGetProcessHandleCount Lib "kernel32" Alias "GetProcessHandleCount" (ByVal hProcess As Long, ByRef lpHandleCount As Long) As Long
Private Declare Function WoGetProcessMemoryInfo Lib "psapi" Alias "GetProcessMemoryInfo" (ByVal hProcess As Long, ByRef ppsmemCounters As Any, ByVal cb As Long) As Long
Private Declare Function WoVirtualQuery Lib "kernel32" Alias "VirtualQuery" (ByVal lpAddress As Long, ByRef lpBuffer As Any, ByVal dwLength As Long) As Long
Private Const WoMEM_COMMIT As Long = &H1000
Private Const WoMEM_FREE As Long = &H10000
Private Declare Function WoGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

' ---- (wheel #1): the fast attach timer --------------------------
' A Win32 timer OWNED BY THIS MODULE. Measured on the live IDE: the watchdog that used
' to be the ONLY driver of WheelScan runs at tmrWatch.Interval = 5000 in steady state
' (measured 2 scans / 10 s), so a code pane opened after the last scan stayed
' un-subclassed for up to ~5 s and the user's wheel went nowhere. The fast attach only
' ever calls Win32 (EnumWindows / EnumChildWindows / GetWindowText / SetWindowLong) -
' NO VBIDE call at all - which is exactly what makes it legal to run from a timer
' callback (a WndProc must never do COM here).
Private Declare Function WoSetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function WoKillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long


' ---- registry access for the settings path --------------------
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



' ---- subclassing trio (only) ----------------------------------
Private Declare Function WoSetWindowLongW Lib "user32" Alias "SetWindowLongW" (ByVal h As Long, ByVal idx As Long, ByVal v As Long) As Long
' R2 WHEEL-FAST: the focus-change channel. WINEVENT_OUTOFCONTEXT posts to THIS thread's
' queue, so the callback is dispatched by the IDE's own message pump on the UI thread and
' can never interleave with a tool call. flags are deliberately 0: WINEVENT_SKIPOWNPROCESS
' (0x0002) would suppress exactly the events we need, because the IDE windows belong to
' OUR process. AddressOf is a direct API argument, the same shape as the wheel subclass
' below (WoSetWindowLongW ... AddressOf WheelWndProc); VB6 rejects p = AddressOf Foo.
Private Declare Function WoSetWinEventHook Lib "user32" Alias "SetWinEventHook" (ByVal evMin As Long, ByVal evMax As Long, ByVal hmod As Long, ByVal cb As Long, ByVal idProc As Long, ByVal idThread As Long, ByVal flags As Long) As Long
Private Declare Function WoUnhookWinEvent Lib "user32" Alias "UnhookWinEvent" (ByVal h As Long) As Long
Private Declare Function WoGetWindowLongW Lib "user32" Alias "GetWindowLongW" (ByVal h As Long, ByVal idx As Long) As Long
Private Declare Function WoCallWindowProcW Lib "user32" Alias "CallWindowProcW" (ByVal prev As Long, ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoDefWindowProcW Lib "user32" Alias "DefWindowProcW" (ByVal h As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Private Declare Function WoGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef pt As WO_POINT) As Long
Private Declare Function WoGetKeyState Lib "user32" Alias "GetKeyState" (ByVal vk As Long) As Integer
' used only to read the code pane's first visible line (see WheelPaneTop).
Private Declare Function WoFindWindowExW Lib "user32" Alias "FindWindowExW" (ByVal parent As Long, ByVal after As Long, ByVal cls As Long, ByVal title As Long) As Long

' ---- window messages ------------------------------------------------------
Private Const WM_NULL As Long = &H0
Private Const WM_CLOSE As Long = &H10
' FIX-11: how long a box may block an in-flight call before the tick stops waiting and
' acts deterministically (safe button -> one-shot WM_CLOSE -> gate release).
Private Const MODAL_WEDGE_DEADLINE_MS As Long = 12000
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

' ---- wheel state (module-level, so it MUST live here: VB6 puts
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
' TRUE only after this add-in has ONCE positively identified a code pane from the
' VBE caption table. It ARMS the empty-table fallback in WheelConsider: at add-in start-up (no
' project, no code pane at all) the table is legitimately empty and hooking anything there would
' gain nothing while making a non-code window the first entry of the receipt's target fields.
Private m_whEverVerified As Boolean
' observability: the last WM_MOUSEWHEEL we saw, what it carried, what we did
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

' ---- final fix: move the pane through the VBIDE object model
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
' (the code-pane caption measurement): where the caption table came from on the last scan
' ("all" | "active_fallback" | "empty") and how many captions it held. Reported in
' ide_state.wheel so a missing hook has an explained cause, not just a symptom.
Private m_whCapsSource As String
Private m_whCapsCount As Long
' (the code-pane caption measurement): which channel resolved the last caption -> CodePane lookup
' ("caption" | "component" | "active_name" | "none").
Private m_whResolveVia As String
' (the code-pane caption measurement): the REAL status/error of the last TopLine read/write, and the
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
' (A/E): per-hook classification. TRUE = the caption table identified this
' window as a code pane on some scan. hooked[] keeps its meaning (code panes
' only); EVERY subclassed window is listed in the new attached[].
Private m_whIsPane(0 To WHEEL_MAX - 1) As Boolean
' (B/E): the fast attach timer plus the per-window decision list of the
' last attach pass (why each VbaWindow was or was not hooked).
Private m_whFastId As Long
Private m_whShutdown As Boolean
Private m_whShutdownCalls As Long
Private m_whSkipAlready As Long
Private m_whHooksInstalled As Long
Private m_woScanned As Long
Private m_woStatMineVisible As Long
Private m_woStatScanTops As Long
Private m_woStatCls As String
Private m_whEnvOff As Boolean
Private m_whAttachReqs As Long
Private m_whAttachChecks As Long
Private m_whAttachHits As Long
Private m_whAttachMisses As Long
Private m_whAttachRetries As Long
Private m_whAttachMaxRetries As Long
Private m_whAttachPending As Boolean
Private m_whAttachDeadline As Long
Private m_whAttachNextAt As Long
Private m_whFocHwnd As Long
Private m_whFocBasis As String
Private m_whFocHooked As Boolean
Private m_whFocusChecks As Long
' R2 WHEEL-FAST: the focus probe is EVENT driven. SetWinEventHook(EVENT_OBJECT_FOCUS) sets
' m_whFocusDirty and WheelTick runs the probe block only on a dirty tick. The 1000 ms
' branch is a correctness BACKSTOP (a dropped event must not strand the probe forever),
' not the mechanism - IdentityWatchTick already uses the same 1000 ms shape. An idle tick
' now performs NO Win32 call and never touches the hook table.
' VB6MCP_WHEEL_FAST=0 restores today's per-tick probe byte for byte.
Private Const WHEEL_EVT_OBJECT_FOCUS As Long = &H8005
Private Const WHEEL_FOCUS_BACKSTOP_MS As Long = 1000
Private m_whEvtHook As Long
Private m_whFocusDirty As Boolean
Private m_whFocusCheckedAt As Long
Private m_whFocusEvents As Long
Private m_whFastSkips As Long
Private m_whHookFallback As Long
Private m_whFastOn As Boolean
' R2 WHEEL-FAST teardown proof: a WinEvent hook holds a FUNCTION POINTER INTO THIS DLL,
' so it MUST be removed before the host unloads us. These counters make the removal
' observable from OUTSIDE (mcp_info.wheel + the logged 'wheel unhook' line) instead of
' trusting that the teardown ran.
Private m_whUnhookCalls As Long
Private m_whUnhookOk As Long
Private m_whReasonRpc As Long
Private m_whReasonFocus As Long
Private m_whReasonArm As Long
Private m_whReasonEvent As Long
Private m_whReasonRetry As Long
Private m_whReasonOther As Long

Private Const WHEEL_ATTACH_GAP_MS As Long = 250
Private Const WHEEL_ATTACH_LADDER_MS As Long = 2000
Private Const WHEEL_ATTACH_MAX_RETRIES As Long = 2
Private Const WHEEL_ATTACH_RETRY_BASE_MS As Long = 300

' Contract id of the environment{} block. A version string, not a number, so a reader can
' compare it literally without knowing this build.
Private Const ENV_SCHEMA_ID As String = "vb6mcp.env/1"
Private m_whFocSeen As Long
Private m_whLastScanAt As Long
Private m_whAttachSuppressed As Long
Private m_pfOn As Boolean
Private m_pfFreq As Double
Private m_pfWindowAt As Double
Private m_pfCpuAt As Double
Private m_pfCpuStart As Double
' patch O: the request METHOD of the last served RPC (truncated ASCII) and the
' tools/list cache counters. Counters are cumulative for the process lifetime.
Private m_pfLastMethod As String
Private m_pfCacheHits As Long
Private m_pfCacheMisses As Long
Private m_pfCacheBytes As Long
' patch P: 4-bucket method counters (lastMethod keeps only the LAST one).
Private m_rpcToolsListN As Long
Private m_rpcMcpInfoN As Long
Private m_rpcOtherN As Long
Private m_rpcExtN As Long
Private m_rpcMcpInfoUs As Double
Private m_rpcToolsListUs As Double
Private m_rpcOtherUs As Double
Private m_rpcExtUs As Double
Private m_rpcGetProjectInfoUs As Double
Private m_findCmdCount As Long
Private m_rpcBucket As Long
Private m_nestedRpcCount As Long
Private m_nestedRpcTickUs As Double
' patch O v2 (H3 instrumentation gap): the OUTBOUND lane was a black hole -
' pollWriteCount only ever counted the overlapped reap, so a synchronously
' completed write contributed zero bytes and the exit byte count was absurd.
Private m_pfOutFrames As Long
Private m_pfOutBytes As Long
Private m_pfWriteSyncBytes As Long
Private m_pfOutReapBytes As Long
Private m_pfLastReqBytes As Long
Private m_pfLastRespBytes As Long
Private m_pfPumpT0 As Double
Private m_pfTicks As Long
Private m_pfTickMs As Long
Private m_pfUs(0 To WO_PF_SEGS - 1) As Double
Private m_pfN(0 To WO_PF_SEGS - 1) As Long
Private m_whFastIn As Boolean
Private m_whFastTicks As Long
Private m_whFastSkipped As Long      ': fast ticks the gate suppressed
Private m_whFastFg As Long           ': foreground hwnd of the last pass
Private m_whFastLastAt As Long
Private m_whAttachRuns As Long
Private m_whDecHwnd(0 To WHEEL_SKIP_MAX - 1) As Long
Private m_whDecTitle(0 To WHEEL_SKIP_MAX - 1) As String
Private m_whDecDecision(0 To WHEEL_SKIP_MAX - 1) As String
Private m_whDecCount As Long
Private Const WHEEL_FAST_MS As Long = 250
' (measured 2026-10-17): WheelAttachPass is a FULL window-tree enumeration.
' It used to run on every 250 ms fast-timer tick, i.e. 240x/min: 3.75 ms x 240 =
' ~900 ms of CPU per 60 s while the IDE was idle - 55% of the add-in's idle cost.
' It is a GAP-CLOSER (a new pane must be hooked before the 5 s watchdog scan sees it),
' so it only needs to run when something may have changed. The gate is the FOREGROUND
' window - the cheap proxy for 'the user moved into a pane' - with a safety net.
Private Const WHEEL_FAST_SAFETY_MS As Long = 5000
' WheelScan and the fast attach BOTH use the WoScan* scratch arrays, and
' WheelScan runs COM (which can pump). One flag keeps them from clobbering each other.
Private m_woScanBusy As Boolean

' ---- armed, one-shot CONFIRM of a modal WE opened ---------------
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

' ---- an earlier fix (defects #10 / #11 / #4): the GENERIC MODAL GUARD ----
' State for the bounded, Cancel-first modal safety net. It lives HERE because VB6
' requires every module-level declaration before the first procedure.
Private m_mgArmed As Boolean
Private m_mgFirstSeenAt As Long
Private m_mgUnresolvable As Boolean
Private m_mgButtons As Object
Private m_mgGlobal As Boolean
Private m_mgAllowRuntimeDebug As Boolean
Private m_mgAllowIdYesNo As Boolean
Private m_mgCategory As String
Private m_mgAction As String
Private m_mgAwaiting As Boolean
Private m_mgDialogs As Object
Private m_mgPre As Object
Private m_mgWedgeAborted As Boolean
Private m_mgWedgeWhy As String
Private m_mgWedgeAt As Long
Private m_mgGateReleased As Boolean
Private m_mgClickedId As Long
Private m_mgCloseSent As Boolean
Private m_mgTickCount As Long
Private m_mgLastTickAction As Object
Private Declare Sub WoKernelSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private m_mgBy As String
Private m_mgAt As Long
Private m_mgDeadline As Long
Private m_mgWaitMs As Long
Private m_mgTries As Long
Private m_mgBusy As Boolean
Private m_mgSeen As Boolean
Private m_mgSeenCount As Long
Private m_mgHwnd As Long
Private m_mgCls As String
Private m_mgTitle As String
Private m_mgBody As String
Private m_mgRisk As String
Private m_mgRiskClass As String
Private m_mgRiskReason As String
Private m_mgClicked As Boolean
Private m_mgClickedText As String
Private m_mgClickedHwnd As Long
Private m_mgMethod As String
Private m_mgDismissed As Boolean
Private m_mgRefused As Long
Private m_mgRefusedReason As String
Private m_mgTimedOut As Boolean
Private m_mgResolvedMs As Long
Private m_mgAllowHostModal As Boolean

' ---- R2 + R12: desktop / interactivity probe (read-only) -------------------
' These are the ONLY new Win32 entry points added for desktop_probe. They READ the
' window-station / desktop / foreground state of THIS process and close the
' input-desktop handle they open. No process memory is read or written, no thread
' is created and nothing is injected - the memory discipline still holds.
Private Declare Function WoGetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
Private Declare Function WoGetUserObjectInformationW Lib "user32" Alias "GetUserObjectInformationW" (ByVal hObj As Long, ByVal nIndex As Long, ByVal pvInfo As Long, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Private Declare Function WoGetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThreadId As Long) As Long
Private Declare Function WoOpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Long, ByVal dwDesiredAccess As Long) As Long
Private Declare Function WoCloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function WoProcessIdToSessionId Lib "kernel32" Alias "ProcessIdToSessionId" (ByVal dwProcessId As Long, ByRef pSessionId As Long) As Long
Private Declare Function WoGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long

Private Const WO_UOI_NAME As Long = 2
Private Const WO_DESKTOP_READOBJECTS As Long = &H1
Private Const WO_DESKTOP_WRITEOBJECTS As Long = &H80
Private Const WO_DESKTOP_SWITCHDESKTOP As Long = &H100

' ---------------------------------------------------------------------
' Host-tick bridge state (the Win32 TIMERPROC facility is defined at the END of
' this module). Declared HERE because VB6 allows only procedures at module level
' after the first procedure, while these are module-level state.
' ---------------------------------------------------------------------
Private m_hostTickId As Long
Private m_hostTickMs As Long
Private m_hostTickSink As Object
Private m_hostTickDepth As Long


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

' the settings path addresses HKCU\Software\VB6MCP only, so callers pass just the
' leaf name; keeping the path here keeps backslashes out of the form's code.
Public Function McpSettingGet(ByVal leaf As String) As String
    McpSettingGet = RegGetString("Software\VB6MCP", leaf)
End Function

Public Function McpSettingSet(ByVal leaf As String, ByVal v As String) As Boolean
    McpSettingSet = RegSetString("Software\VB6MCP", leaf, v)
End Function

Private Function WoText(ByVal h As Long) As String
    ' FIX-11: ANSI bytes -> ONE CP_ACP decode. The old UTF-16 read turned the IDE's
    ' owner-draw button captions into mojibake (0x6D60...), which defeated every caption
    ' whitelist; standard controls still come out correct because Windows does the
    ' Unicode->ANSI conversion for GetWindowTextA itself.
    Dim b As String, n As Long, w As String, m As Long
    On Error Resume Next
    WoText = ""
    b = String$(512, vbNullChar)
    n = WoGetWindowTextA(h, StrPtr(b), 512)
    If n <= 0 Then Err.Clear : Exit Function
    b = Left$(b, n)
    w = String$(1024, vbNullChar)
    m = WoMultiByteToWideChar(0&, 0&, StrPtr(b), n, StrPtr(w), 1024)
    If m > 0 Then
        WoText = Left$(w, m)
    Else
        WoText = b
    End If
    Err.Clear
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
' (1): read_active_dialogs
' Enumerate the ACTIVE dialogs of THIS process (default class #32770) and
' report class / title / geometry / owner / whether the OWNER is DISABLED (the
' only reliable "this box is blocking its parent" signal - "a window exists" is
' not enough) plus the child controls and the clickable button list.
' READ-ONLY: it enumerates and reads text; it never sends anything.
' =====================================================================

Private Function WoCollect(ByVal h As Long, ByVal lp As Long) As Long
    m_woScanned = m_woScanned + 1
    If m_woCount < WO_MAX Then
        m_woHwnd(m_woCount) = h
        m_woCount = m_woCount + 1
    End If
    WoCollect = 1
End Function

' Owning THREAD of a window (the dialog report needs it: a box owned by another thread
' of this process is exactly the case read_active_dialogs must still describe).
' The FOCUSED window of THIS thread, O(1) and INDEPENDENT of who owns the foreground.
' GetFocus alone is NOT usable: it answers 0 unless the calling thread's queue owns the
' keyboard focus, so a background IDE would look unfocused forever and the attach probe
' would silently degrade to 'never sees an unhooked window'. GetGUIThreadInfo reports
' the thread's own hwndFocus whatever the foreground is; the foreground window's thread
' is the FALLBACK, and the receipt names which basis answered (focBasis).
Private Function WoGetThreadFocus() As Long
    On Error Resume Next
    Dim gti As WO_GUITHREADINFO
    Dim fg As Long
    WoGetThreadFocus = 0
    m_whFocBasis = "none"
    gti.cbSize = WO_GTI_SIZE
    If WoGetGUIThreadInfo(WoGetCurrentThreadId(), gti) <> 0 Then
        If gti.hwndFocus <> 0 Then
            m_whFocBasis = "guithreadinfo"
            WoGetThreadFocus = gti.hwndFocus
            Err.Clear
            Exit Function
        End If
    End If
    fg = WoGetForegroundWindow()
    If fg <> 0 Then
        gti.cbSize = WO_GTI_SIZE
        If WoGetGUIThreadInfo(WoThreadOf(fg), gti) <> 0 Then
            If gti.hwndFocus <> 0 Then
                m_whFocBasis = "foreground_thread"
                WoGetThreadFocus = gti.hwndFocus
            End If
        End If
    End If
    Err.Clear
End Function

Private Function WoThreadOf(ByVal h As Long) As Long
    Dim p As Long
    p = 0
    WoThreadOf = WoGetWindowThreadProcessId(h, p)
End Function

Private Sub WoScanTops()
    m_woScanned = 0
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
    ' Snapshot what the top-level scan saw BEFORE any child enumeration overwrites the
    ' shared window buffer: this is what makes an empty answer explainable.
    m_woStatScanTops = m_woScanned
    m_woStatMineVisible = 0
    m_woStatCls = ""
    Dim stI As Long
    Dim stH As Long
    Dim stC As String
    For stI = 0 To nTops - 1
        stH = tops(stI)
        If stH <> 0 Then
            If WoPidOf(stH) = WoGetCurrentProcessId() Then
                If WoIsWindowVisible(stH) <> 0 Then
                    m_woStatMineVisible = m_woStatMineVisible + 1
                    stC = LCase$(WoCls(stH))
                    If InStr(1, m_woStatCls, "[" & stC & "]") = 0 Then
                        m_woStatCls = m_woStatCls & "[" & stC & "]"
                    End If
                End If
            End If
        End If
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
                        d("threadId") = WoThreadOf(h)
                        d("textReadable") = (Len(CStr(d("title"))) > 0)
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
    ' A dialog whose body could not be read is still REPORTED (hwnd/cls/thread), not
    ' dropped: dialogTextUnreadable says the text is missing, count is never a silent 0.
    Dim unreadable As Long
    Dim ui As Long
    unreadable = 0
    For ui = 1 To arr.Count
        If Not CBool(arr.Item(ui)("textReadable")) Then unreadable = unreadable + 1
    Next
    If arr.Count > 0 Then
        root("dialogTextUnreadable") = (unreadable = arr.Count)
        root("dialogsWithoutText") = unreadable
    End If
    root("scanTops") = m_woStatScanTops
    root("scanOk") = (m_woStatScanTops > 0)
    root("processVisibleTopWindows") = m_woStatMineVisible
    root("processWindowClasses") = m_woStatCls
    root("scanBasis") = "scanTops counts EVERY top-level window the enumeration walked; processVisibleTopWindows counts the visible ones owned by this process; processWindowClasses lists their classes, so an unmatched dialog is visible instead of a silent zero"
    If m_woStatScanTops = 0 Then
        root("dialogScanBlocked") = True
        root("scanBlockedReason") = "the top-level enumeration returned no windows at all"
    Else
        root("dialogScanBlocked") = False
    End If
    Set root("dialogs") = arr
    If arr.Count = 0 Then
        root("note") = "no visible top-level window of this process matches cls=" & cls & ". A box that is up appears here; ownerDisabled=true on it is the signal that it owns (blocks) the IDE"
    End If
    ReadActiveDialogsJson = JsonStringify(root)
End Function

' =====================================================================
' (2): close_active_dialogs
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
' (S point 3): cancelling cannot START the host-modal operation, so a
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

' "No" is NOT on the default whitelist (for a "save changes?" box it
' discards work) but it is never DESTRUCTIVE, so an EXPLICIT button="No" is the escape
' hatch the body guard suggests. Only the exact captions No / Chinese No qualify.
Private Function WoButtonExplicitNo(ByVal norm As String) As Boolean
    WoButtonExplicitNo = WoExactHit("NO|" & ChrW$(&H5426), norm)
End Function


' =====================================================================
' (39) DIALOG-ID-FIRST: control ID primary, caption fallback.
' Measured ID tables (fix/TASK_G5_IDFIRST_20261024/DESIGN.md, G5/W41):
'   runtime_error : continue=4802 (measured enabled=false), end=4800,
'                   debug=4801, help=4902, body Static=4803
'   save_confirm  : yes=6, no=7, cancel=2, Static=65535, ListBox=5155
'   eject_refused : caption OK is id=2 (IDCANCEL), help=9
'   load_fail     : yes=6, no=7, help=9
' The table is keyed by DIALOG KIND on purpose: the SAME number means
' different things in different boxes (id=2 is Cancel in the save box but
' the "OK" button of the eject-refused box), so an ID is never reused
' across kinds. Text is used ONLY to (a) pick the kind domain and (b) run
' the SECONDARY cross-check that can refuse a click (id_text_conflict).
' =====================================================================
Public Function WoDlgIdTableVersion() As String
    ' measured ID-table revision (G5 findings + W41 runtime-error box, 2026-10-23)
    WoDlgIdTableVersion = "g5r1-2026-10-23"
End Function

Public Function WoDlgIdFirstOn() As Boolean
    ' default ON; VB6MCP_DLG_ID_FIRST=0/off/false/no forces the legacy
    ' caption-only path (rollback switch, no rebuild of the ID tables needed)
    Dim v As String
    WoDlgIdFirstOn = True
    On Error Resume Next
    v = LCase$(Trim$(Environ("VB6MCP_DLG_ID_FIRST")))
    If Len(v) = 0 Then Exit Function
    If v = "0" Or v = "off" Or v = "false" Or v = "no" Then WoDlgIdFirstOn = False
    Err.Clear
End Function

Public Function WoDlgKindOf(ByVal d As Object) As String
    ' text -> DOMAIN only (which ID table applies). The button is still chosen by ID.
    WoDlgKindOf = "unknown"
    On Error Resume Next
    If d Is Nothing Then Exit Function
    Dim body As String, ttl As String
    body = UCase$(WoMgBodyOf(d))
    ttl = ""
    If d.Exists("title") Then ttl = UCase$(CStr(d("title")))
    If InStr(1, body, "RUN-TIME ERROR") > 0 Then WoDlgKindOf = "runtime_error": Err.Clear: Exit Function
    If InStr(1, body, "RUNTIME ERROR") > 0 Then WoDlgKindOf = "runtime_error": Err.Clear: Exit Function
    If InStr(1, body, ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then WoDlgKindOf = "runtime_error": Err.Clear: Exit Function
    If InStr(1, body, ChrW$(&H5B9E) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then WoDlgKindOf = "runtime_error": Err.Clear: Exit Function
    If InStr(1, ttl, "SAVE") > 0 Then WoDlgKindOf = "save_confirm": Err.Clear: Exit Function
    If InStr(1, ttl, ChrW$(&H4FDD) & ChrW$(&H5B58)) > 0 Then WoDlgKindOf = "save_confirm": Err.Clear: Exit Function
    If InStr(1, body, ChrW$(&H53E6) & ChrW$(&H5B58)) > 0 Then WoDlgKindOf = "save_confirm": Err.Clear: Exit Function
    If InStr(1, body, "CANNOT EXIT") > 0 Then WoDlgKindOf = "eject_refused": Err.Clear: Exit Function
    If InStr(1, body, ChrW$(&H6B64) & ChrW$(&H65F6) & ChrW$(&H4E0D) & ChrW$(&H80FD) & ChrW$(&H9000) & ChrW$(&H51FA)) > 0 Then WoDlgKindOf = "eject_refused": Err.Clear: Exit Function
    If InStr(1, body, "CANNOT LOAD") > 0 Then WoDlgKindOf = "load_fail": Err.Clear: Exit Function
    If InStr(1, body, ChrW$(&H4E0D) & ChrW$(&H80FD) & ChrW$(&H52A0) & ChrW$(&H8F7D)) > 0 Then WoDlgKindOf = "load_fail": Err.Clear: Exit Function
    Err.Clear
End Function

Public Function WoDlgIdSemantic(ByVal kind As String, ByVal cid As Long) As String
    ' "" = this ID is NOT in the domain's table (caller must fall back to text).
    WoDlgIdSemantic = ""
    Select Case LCase$(kind)
        Case "runtime_error"
            Select Case cid
                Case 4802&: WoDlgIdSemantic = "continue"
                Case 4800&: WoDlgIdSemantic = "end"
                Case 4801&: WoDlgIdSemantic = "debug"
                Case 4902&: WoDlgIdSemantic = "help"
                Case 4803&: WoDlgIdSemantic = "body"
            End Select
        Case "save_confirm"
            Select Case cid
                Case 6&: WoDlgIdSemantic = "yes"
                Case 7&: WoDlgIdSemantic = "no"
                Case 2&: WoDlgIdSemantic = "cancel"
                Case 65535&: WoDlgIdSemantic = "body"
                Case 5155&: WoDlgIdSemantic = "filelist"
            End Select
        Case "eject_refused"
            Select Case cid
                Case 2&: WoDlgIdSemantic = "ok"
                Case 9&: WoDlgIdSemantic = "help"
            End Select
        Case "load_fail"
            Select Case cid
                Case 6&: WoDlgIdSemantic = "yes"
                Case 7&: WoDlgIdSemantic = "no"
                Case 9&: WoDlgIdSemantic = "help"
            End Select
    End Select
End Function

Public Function WoDlgIdSemanticClickable(ByVal kind As String, ByVal sem As String) As Boolean
    ' SAFETY: continue/end are NEVER clicked (measured End 4800 enabled; W41
    ' dismiss must stay Debug-only). help/body/filelist are not actions.
    WoDlgIdSemanticClickable = False
    Select Case LCase$(sem)
        Case "yes", "no", "cancel", "ok", "debug": WoDlgIdSemanticClickable = True
    End Select
End Function

Private Function WoDlgSemNeedContinue() As String
    ' no-op helper kept out; placeholder never used
    WoDlgSemNeedContinue = ""
End Function

Public Function WoDlgWantSemantic(ByVal kind As String, ByVal want As String) As String
    ' WANT caption (already accelerator-stripped + uppercased by the caller) -> semantic.
    ' "@none" = this kind must NOT be clicked through the ID path at all.
    WoDlgWantSemantic = ""
    Dim w As String
    w = UCase$(Trim$(Replace(want, "&", "")))
    Select Case LCase$(kind)
        Case "runtime_error"
            ' CloseActiveDialogs must never press Continue/End/Debug/Help by ID.
            WoDlgWantSemantic = "@none"
            Exit Function
        Case "save_confirm"
            If w = "YES" Or w = ChrW$(&H662F) Then WoDlgWantSemantic = "yes": Exit Function
            If w = "NO" Or w = ChrW$(&H5426) Then WoDlgWantSemantic = "no": Exit Function
            If w = "CANCEL" Or w = ChrW$(&H53D6) & ChrW$(&H6D88) Or w = "CLOSE" Or w = ChrW$(&H5173) & ChrW$(&H95ED) Then WoDlgWantSemantic = "cancel": Exit Function
        Case "eject_refused", "load_fail"
            If w = "OK" Or w = ChrW$(&H786E) & ChrW$(&H5B9A) Then WoDlgWantSemantic = "ok": Exit Function
            If w = "CANCEL" Or w = ChrW$(&H53D6) & ChrW$(&H6D88) Then WoDlgWantSemantic = "cancel": Exit Function
    End Select
End Function

Public Function WoDlgIdResolve(ByVal kind As String, ByVal cid As Long, ByVal cls As String, _
                               ByVal enabled As Boolean, ByVal text As String, ByVal want As String) As Object
    ' PURE decision (no I/O): the ONE place that says whether a control ID may be
    ' clicked in this domain. Keys: ok, idBasis, controlId, semantics, wantSemantic,
    ' reason, textHit, idTableVersion, clickable.
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    r("ok") = False
    r("idBasis") = "none"
    r("controlId") = cid
    r("semantics") = ""
    r("wantSemantic") = ""
    r("reason") = ""
    r("textHit") = ""
    r("clickable") = False
    r("idTableVersion") = WoDlgIdTableVersion()
    Set WoDlgIdResolve = r
    On Error Resume Next
    If Not WoDlgIdFirstOn() Then r("reason") = "id_first_disabled": Err.Clear: Exit Function
    If Len(kind) = 0 Then r("reason") = "kind_unknown": Err.Clear: Exit Function
    If LCase$(kind) = "unknown" Then r("reason") = "kind_unknown": Err.Clear: Exit Function
    If InStr(1, LCase$(cls), "button", vbTextCompare) = 0 Then r("reason") = "not_a_button": Err.Clear: Exit Function
    Dim sem As String
    sem = WoDlgIdSemantic(kind, cid)
    r("semantics") = sem
    If Len(sem) = 0 Then r("reason") = "id_not_in_table": Err.Clear: Exit Function
    If Not enabled Then r("reason") = "button_disabled": Err.Clear: Exit Function
    r("clickable") = WoDlgIdSemanticClickable(kind, sem)
    If Not CBool(r("clickable")) Then r("reason") = "never_click_dangerous": Err.Clear: Exit Function
    If Len(want) > 0 Then
        If want = "@none" Then r("reason") = "kind_not_clickable_by_policy": Err.Clear: Exit Function
        r("wantSemantic") = want
        If StrComp(want, sem, vbTextCompare) <> 0 Then r("reason") = "semantic_mismatch": Err.Clear: Exit Function
    End If
    ' SECONDARY text cross-check: a text that contradicts the ID semantics is a
    ' hard refusal (the ID may come from a language/build whose table we do not know).
    Dim hit As String
    If WoDlgIdTextConflict(sem, text, hit) Then
        r("reason") = "id_text_conflict"
        r("textHit") = hit
        Err.Clear
        Exit Function
    End If
    r("ok") = True
    r("idBasis") = "control_id"
    Err.Clear
End Function

Private Function WoDlgIdTextConflict(ByVal sem As String, ByVal text As String, ByRef hit As String) As Boolean
    ' True = the caption says something a different semantic than the ID table.
    ' Empty caption is NOT a conflict (some custom controls have no text).
    WoDlgIdTextConflict = False
    hit = ""
    Dim n As String
    n = UCase$(Trim$(Replace(text, "&", "")))
    If Len(n) = 0 Then Exit Function
    Dim p As Long
    p = InStr(n, "(")
    If p > 0 Then n = Trim$(Left$(n, p - 1))
    Do While Len(n) > 0
        If Right$(n, 1) = "." Then n = Left$(n, Len(n) - 1) Else Exit Do
    Loop
    n = Trim$(n)
    If Len(n) = 0 Then Exit Function
    Select Case LCase$(sem)
        Case "debug"
            If InStr(1, n, "DEBUG") = 0 Then
                If InStr(1, text, ChrW$(&H8C03) & ChrW$(&H8BD5)) = 0 Then hit = "expected_debug_caption": WoDlgIdTextConflict = True
            End If
        Case "end"
            If n <> "END" Then
                If InStr(1, text, ChrW$(&H7ED3) & ChrW$(&H675F)) = 0 Then hit = "expected_end_caption": WoDlgIdTextConflict = True
            End If
        Case "continue"
            If InStr(1, n, "CONTINUE") = 0 Then
                If InStr(1, text, ChrW$(&H7EE7) & ChrW$(&H7EED)) = 0 Then hit = "expected_continue_caption": WoDlgIdTextConflict = True
            End If
        Case "help"
            If n <> "HELP" Then
                If InStr(1, text, ChrW$(&H5E2E) & ChrW$(&H52A9)) = 0 Then hit = "expected_help_caption": WoDlgIdTextConflict = True
            End If
        Case "cancel"
            If Not WoButtonCancelLike(n) Then hit = "expected_cancel_caption": WoDlgIdTextConflict = True
        Case "yes"
            If n <> "YES" Then
                If InStr(1, text, ChrW$(&H662F)) = 0 Then hit = "expected_yes_caption": WoDlgIdTextConflict = True
            End If
        Case "no"
            If n <> "NO" Then
                If InStr(1, text, ChrW$(&H5426)) = 0 Then hit = "expected_no_caption": WoDlgIdTextConflict = True
            End If
    End Select
End Function

' destructive-action patterns read from the dialog BODY (title +
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

' the dialog BODY as the guard sees it = title + every child
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
' the dialog-interlock classifier.
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

' public, read-only. The worst host-modal risk among the dialogs now
' up, plus per-dialog detail as JSON. Returns the dialog count. Used by
' make_project before it decides whether to confirm the Make box.
' cheap, read-only: is a VISIBLE top-level #32770 of THIS process up right now?
' Used by Util.PumpGateBlocked; it never clicks, closes or messages anything.
Public Function HasVisibleHostDialog(ByRef titleOut As String) As Boolean
    On Error Resume Next
    HasVisibleHostDialog = False
    titleOut = ""
    Dim n As Long, i As Long
    WoScanTops
    n = m_woCount
    If n > WO_MAX Then n = WO_MAX
    If n <= 0 Then Exit Function
    Dim pid As Long
    pid = WoGetCurrentProcessId()
    For i = 0 To n - 1
        Dim h As Long
        h = m_woHwnd(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                If WoCls(h) = "#32770" Then
                    If WoIsWindowVisible(h) <> 0 Then
                        titleOut = WoText(h)
                        HasVisibleHostDialog = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function

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

' the visible text of every active dialog of this process,
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

' R-1021-05: TYPED (JSON-free) view of the visible dialogs of THIS process: exactly the
' structure ReadActiveDialogsJson serialises, handed over as a Collection of Dictionary so
' the runtime-error capture can read a body and click a button from ONE enumerating scan.
' Read-only in the strict sense: it enumerates windows and reads their text; it never
' sends a message.
Public Function DialogObjectsByClass(ByVal cls As String, ByVal withChildren As Boolean, ByVal maxN As Long) As Object
    On Error Resume Next
    Dim out As Object
    Set out = New Collection
    If Len(cls) = 0 Then cls = "#32770"
    Dim ds As Object
    Set ds = WoDialogsJson(cls, 0, withChildren)
    If ds Is Nothing Then
        Set DialogObjectsByClass = out
        Exit Function
    End If
    Dim i As Long
    For i = 1 To ds.Count
        If maxN > 0 Then
            If out.Count >= maxN Then Exit For
        End If
        out.Add ds.Item(i)
    Next
    Set DialogObjectsByClass = out
    Err.Clear
End Function


' best-effort text of a VB6 compile-error box. The box body
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
    root("policy") = "clicks ONLY buttons whose caption is on the safe whitelist (OK/Yes/Cancel/Close/Continue/Ignore + Chinese twins) and never on the danger denylist; the click is the BM_CLICK window message (a WM_COMMAND fallback is sent only if the button still exists). A second, independent guard: a dialog that would start a LONG host-modal operation (VB6's Make Project box) is reported as host_modal_risk=known and is NOT clicked unless allow_host_modal=true."
    ' (discipline 98): the same key set on EVERY path. Without these
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
        ' ---- classify BEFORE anything is sent --------------
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
        ' FIX-9: the SAME selector as the modal guard: caption whitelist first, then
        ' standard control IDs IDCANCEL=2 -> IDOK=1 -> IDABORT=3. IDYES/IDNO are refused
        ' (no opt-in here), and WoMgIdFallbackAllowed refuses End/Continue/runtime-error
        ' boxes. This is what stops F10: a required '&Open' button is not on the caption
        ' whitelist, and skipping it used to leave the modal over the IDE.
        Dim fbHwnd As Long
        Dim fbId As Long
        Dim fbPass As Long
        Dim pickedBasis As String
        Dim pickedControlId As Long
        fbHwnd = 0
        fbId = 0
        pickedBasis = "caption_fallback"
        pickedControlId = 0
        Dim dlgKind As String
        dlgKind = WoDlgKindOf(d)
        Dim wantSem As String
        wantSem = WoDlgWantSemantic(dlgKind, want)
        If Not btns Is Nothing Then
            If WoDlgIdFirstOn() And Len(dlgKind) > 0 And wantSem <> "@none" Then
                If Len(want) = 0 Or Len(wantSem) > 0 Then
                If dlgKind <> "unknown" Then
                    For j = 1 To btns.Count
                        Dim idBtn As Object
                        Set idBtn = btns.Item(j)
                        If CBool(idBtn("enabled")) Then
                            Dim idRes As Object
                            Set idRes = WoDlgIdResolve(dlgKind, CLng(idBtn("id")), CStr(idBtn("cls")), CBool(idBtn("enabled")), CStr(idBtn("text")), wantSem)
                            If Not idRes Is Nothing Then
                                If CBool(idRes("ok")) Then
                                    fbHwnd = CLng(idBtn("hwnd"))
                                    fbId = CLng(idBtn("id"))
                                    pickedBasis = "control_id"
                                    pickedControlId = fbId
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                End If
                End If
            End If
            If fbHwnd = 0 And wantSem <> "@none" Then
                If WoMgIdFallbackAllowed(d, btns) Then
                    For fbPass = 1 To 3
                        Dim fbWant As Long
                        If fbPass = 1 Then
                            fbWant = 2
                        ElseIf fbPass = 2 Then
                            fbWant = 1
                        Else
                            fbWant = 3
                        End If
                        For j = 1 To btns.Count
                            Dim fbB As Object
                            Set fbB = btns.Item(j)
                            If CBool(fbB("enabled")) Then
                                If CLng(fbB("id")) = fbWant Then
                                    fbHwnd = CLng(fbB("hwnd"))
                                    fbId = fbWant
                                    pickedBasis = "control_id"
                                    pickedControlId = fbWant
                                    Exit For
                                End If
                            End If
                        Next
                        If fbHwnd <> 0 Then Exit For
                    Next
                End If
            End If
        End If
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
                ' skipped[] carries the SAME key set on every
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
                    ElseIf dlgRisk = "known" And Not allowHostModal And Not WoButtonCancelLike(normA) _
                           And Not (fbHwnd <> 0 And fbId = 2 And CLng(b("hwnd")) = fbHwnd) Then
                        ' this is the click that cost the user the whole
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
                    ElseIf (Not WoButtonAllow(t, norm)) And Not (Len(want) > 0 And WoButtonExplicitNo(norm)) _
                           And Not (fbHwnd <> 0 And CLng(b("hwnd")) = fbHwnd) Then
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
                        ck("clickedCaption") = t
                        ck("clickedHwnd") = b("hwnd")
                        ck("clickedId") = b("id")
                        ck("idBasis") = pickedBasis
                        ck("controlId") = CLng(b("id"))
                        ck("idTableVersion") = WoDlgIdTableVersion()
                        If fbHwnd <> 0 Then
                            If CLng(b("hwnd")) = fbHwnd Then
                                ck("standardIdFallback") = True
                                ck("fallbackId") = fbId
                            End If
                        End If
                        ' ONE key set on BOTH the dry_run and the
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
        Dim remD As Object
        Set remD = after.Item(i)
        Dim remRec As Object
        Set remRec = WoRecoveryForDialog(remD)
        If Not remRec Is Nothing Then Set remD("recovery") = remRec
        remaining.Add remD
    Next
    If remaining.Count > 0 Then
        Set root("recovery") = WoRecoveryForDialog(after.Item(1))
    End If
    root("dialogsBefore") = before.Count
    root("clickedCount") = clicked.Count
    root("skippedCount") = skipped.Count
    Set root("clicked") = clicked
    Set root("skipped") = skipped
    root("dialogsAfter") = remaining.Count
    Set root("remaining") = remaining
    ' ---- the host-modal verdict, on EVERY path ---------
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
' the ARMED, ONE-SHOT confirm used by the IDE-native form route.
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

    ' an earlier fix : when the GENERIC MODAL GUARD is armed it OWNS the dialog
    ' decision, because this one-shot confirm only ever presses a CONFIRM caption
    ' (Open / OK / the Chinese twins). On VB6's own Make Project box the enumeration
    ' order is Open(read-only,disabled), OK(id1), Cancel(id2)... so pressing "the first
    ' confirm caption" STARTS the compile inside the host's modal loop - defect #4.
    ' The guard is Cancel-first and never does that.
    If m_mgArmed Then Exit Sub

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


' =====================================================================
' an earlier fix (defects #10 / #11 / #4): the GENERIC MODAL GUARD
'
' WHY. Three unrelated tools were measured to leave a MODAL #32770 box up and
' then never return:
'   #11 native_breakpoint{toggle, <a line VB6 refuses>} -> "a breakpoint cannot be
'       set on this line" (the Debug menu item blocks inside the modal loop);
'   #10 vba6_probe{exec_probe} on a raising statement -> VB6's run-time error box;
'   #4  make_project's failure-fallback File>Make -> VB6's own Make Project box.
' In all three the add-in is single-threaded and the tool call is SUSPENDED
' INSIDE the host's modal message loop, so no code after the failing call can ever
' dismiss the box. The only channel that still runs is a message that modal loop
' itself dispatches: the watchdog tick (ServerForm.tmrWatch -> WheelTick), which is
' the same channel WoConfirmTick already proved works.
'
' WHAT. Arm the guard around ONE risky call. While armed, every pass enumerates the
' visible #32770 windows of THIS process that BLOCK their owner (ownerDisabled=true)
' and presses at most ONE button:
'   * ALWAYS preferred: an exact Cancel-like caption (Cancel / Close / Ignore) or an
'     explicit No. Cancelling cannot START the operation, so it is safe even on a box
'     this call did not open.
'   * otherwise, ONLY when the box is NOT host-modal risk=known: the first ENABLED
'     button whose caption is on the SAFE whitelist close_active_dialogs already uses
'     (WoButtonAllow), provided the dialog BODY names no destructive action.
'   * otherwise: NOTHING is sent, and the refusal is reported.
' VB6's own Make Project box enumerates its buttons Open-as-read-only(disabled),
' OK(id1), Cancel(id2), Help, Options. "The first safe button" would therefore press
' OK and start a compile inside that same modal loop - which IS defect #4. The
' Cancel-first rule is what makes this guard safe.
'
' HONESTY. This is a bounded SAFETY NET, not a proof:
'   * a box that never appears inside the budget leaves modalDetected=false and
'     recoveredBy="not_needed" - that is NOT a claim that no box existed;
'   * a box that appears but cannot be pressed leaves modalDetected=true with an
'     empty clickedButton and recoveredBy="modal_guard_timeout";
'   * nothing is ever swallowed silently - the receipt is on EVERY path of every
'     tool that arms this guard.
' =====================================================================

' FIX-7 P0: the GLOBAL guard is armed once per TOOL CALL by the dispatcher. Site arms turn
' into no-ops while it is on, so an inner site cannot clobber the outer deadline/hwnd, and a
' site disarm cannot kill the outer guard mid-call.
Public Sub WoModalGuardArmGlobal(ByVal byTool As String, ByVal waitMs As Long, ByVal allowRuntimeDebug As Boolean)
    m_mgGlobal = True
    WoModalGuardArm byTool, waitMs, False, allowRuntimeDebug, False
End Sub

Public Sub WoModalGuardDisarmGlobal()
    m_mgGlobal = False
    WoModalGuardDisarm
End Sub

Public Sub WoModalGuardSiteArm(ByVal byTool As String, ByVal waitMs As Long, Optional ByVal allowHostModal As Boolean = False)
    If m_mgGlobal Then Exit Sub
    WoModalGuardArm byTool, waitMs, allowHostModal
End Sub

Public Sub WoModalGuardSiteDisarm()
    If m_mgGlobal Then Exit Sub
    WoModalGuardDisarm
End Sub

Public Function WoModalGuardArmed() As Boolean
    WoModalGuardArmed = m_mgArmed
End Function

Public Sub WoModalGuardArm(ByVal byTool As String, ByVal waitMs As Long, Optional ByVal allowHostModal As Boolean = False, Optional ByVal allowRuntimeDebug As Boolean = False, Optional ByVal allowIdYesNo As Boolean = False)
    On Error Resume Next
    m_mgBy = byTool
    m_mgAt = WoGetTickCount()
    If waitMs < 1000 Then waitMs = 1000
    If waitMs > 120000 Then waitMs = 120000
    m_mgWaitMs = waitMs
    m_mgDeadline = m_mgAt + waitMs
    m_mgTries = 0
    m_mgBusy = False
    m_mgSeen = False
    m_mgSeenCount = 0
    m_mgHwnd = 0
    m_mgCls = ""
    m_mgTitle = ""
    m_mgBody = ""
    m_mgRisk = ""
    m_mgRiskClass = ""
    m_mgRiskReason = ""
    m_mgClicked = False
    m_mgClickedText = ""
    m_mgClickedHwnd = 0
    m_mgMethod = ""
    m_mgDismissed = False
    m_mgRefused = 0
    m_mgRefusedReason = ""
    m_mgTimedOut = False
    m_mgResolvedMs = 0
    m_mgAllowHostModal = allowHostModal
    m_mgFirstSeenAt = 0
    m_mgUnresolvable = False
    Set m_mgButtons = Nothing
    m_mgAllowRuntimeDebug = allowRuntimeDebug
    m_mgAllowIdYesNo = allowIdYesNo
    m_mgCategory = ""
    m_mgAction = ""
    m_mgAwaiting = False
    m_mgWedgeAborted = False
    m_mgWedgeWhy = ""
    m_mgWedgeAt = 0
    m_mgGateReleased = False
    m_mgClickedId = 0
    m_mgCloseSent = False
    m_mgTickCount = 0
    Set m_mgLastTickAction = Nothing
    Set m_mgDialogs = New Collection
    ' FIX-9: remember which #32770 boxes already existed BEFORE this call. They are still
    ' handled by the guard (otherwise the next call would keep being blocked by them), but
    ' every dialogs[] entry says preExisting=true so the ledger never pretends they were
    ' caused by this call.
    Set m_mgPre = New Collection
    If WoCfmDialogs() > 0 Then
        Dim pIdx As Long
        For pIdx = 0 To m_cfmScanN - 1
            m_mgPre.Add m_cfmScan(pIdx)
        Next
    End If
    m_mgArmed = True
    Err.Clear
End Sub

' Disarm. When the budget ran out with a box SEEN but nothing pressed, that is
' recorded as a timeout instead of being left ambiguous.
Public Sub WoModalGuardDisarm()
    On Error Resume Next
    If m_mgArmed Then
        If m_mgSeen And Not m_mgClicked Then
            If WoGetTickCount() >= m_mgDeadline Then m_mgTimedOut = True
        End If
    End If
    ' an earlier fix: the LAST chance to OBSERVE the result of a click. BM_CLICK is sent
    ' synchronously, but a MODAL dialog is torn down by the host's own modal loop when it
    ' regains control - i.e. AFTER the callback that sent the click has returned. On the
    ' first build of this card that made modalDismissed stay False and the receipt say
    ' "unknown ... still open at the last check" even though the box was gone. One
    ' DoEvents lets that teardown run, so "dismissed" can be a FACT here; when the box is
    ' genuinely still there it is reported as still there rather than guessed.
    If m_mgClicked Then
        ' same gate - observing the dismissal must never dispatch a POSTED host
        ' command. Without the pump the click is still sent; only the OBSERVATION is lost.
        If CanPump() Then
            m_pfPumpT0 = WoProfNow()
            DoEvents
            WoPumpAdd WoProfNow() - m_pfPumpT0
        End If
        If WoIsWindow(m_mgHwnd) = 0 Then m_mgDismissed = True
    End If
    m_mgArmed = False
    Err.Clear
End Sub

' PURE READ: TRUE while the guard still has work to do. It never changes state, so a
' caller can pump on it and then disarm explicitly.
Public Function WoModalGuardPending() As Boolean
    On Error Resume Next
    If Not m_mgArmed Then Exit Function
    If m_mgRefused > 0 Then Exit Function
    WoModalGuardPending = (WoGetTickCount() < m_mgDeadline)
    Err.Clear
End Function

' PURE READ: TRUE once a BLOCKING dialog has been SEEN during this arm cycle. A caller
' uses it to decide whether the full budget is still worth waiting for: with no box in
' sight the pump only needs a short grace, and waiting the whole budget on every healthy
' call would be a needless 15 s stall.
Public Function WoModalGuardSawBox() As Boolean
    WoModalGuardSawBox = m_mgSeen
End Function

' The dialog BODY as this guard sees it: title plus every child control's text.
Private Function WoMgBodyOf(ByVal d As Object) As String
    On Error Resume Next
    Dim s As String
    s = CStr(d("title"))
    Dim ctrls As Object
    Set ctrls = d("controls")
    If Not ctrls Is Nothing Then
        Dim i As Long, ce As Object
        For i = 1 To ctrls.Count
            Set ce = ctrls.Item(i)
            s = s & " | " & CStr(ce("text"))
        Next
    End If
    If Len(s) > 800 Then s = Left$(s, 800) & "...[truncated]"
    WoMgBodyOf = s
    Err.Clear
End Function

' Pick the ONE button this guard is allowed to press in dialog d.
' Pass 1 = cancel-like (always preferred; safe on a box this call did not open).
' Pass 2 = the shared SAFE whitelist; the CALLER still refuses this one when the
' dialog is host-modal risk=known, because on VB6's Make box it is "OK" = compile.
Private Function WoMgPickButton(ByVal d As Object, ByRef outText As String, ByRef outWhy As String, ByRef outCancel As Boolean, ByRef outId As Long) As Long
    On Error Resume Next
    outText = ""
    outWhy = "no_button"
    outCancel = False
    outId = 0
    Dim btns As Object
    Set btns = btnsOf(d)
    If btns Is Nothing Then Exit Function
    Dim j As Long, b As Object, t As String, norm As String
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        t = CStr(b("text"))
        norm = UCase$(Trim$(Replace(t, "&", "")))
        If CBool(b("enabled")) Then
            If WoButtonCancelLike(norm) Or WoButtonExplicitNo(norm) Then
                outText = t
                outCancel = True
                outWhy = "cancel_like_caption"
                outId = CLng(b("id"))
                WoMgPickButton = CLng(b("hwnd"))
                Exit Function
            End If
        End If
    Next
    Dim n2 As String
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        t = CStr(b("text"))
        n2 = ""
        If CBool(b("enabled")) Then
            If WoButtonAllow(t, n2) Then
                outText = t
                outCancel = False
                outWhy = "safe_whitelist_caption"
                outId = CLng(b("id"))
                WoMgPickButton = CLng(b("hwnd"))
                Exit Function
            End If
        End If
    Next
    ' pass 3 (FIX-7 rev): STANDARD CONTROL IDs. Some boxes have localized captions
    ' the whitelist has never seen yet carry the standard Win32 IDs (IDCANCEL=2,
    ' IDOK=1, IDABORT=3). Only used when the box is provably NOT an End/Continue-class
    ' (runtime-error) dialog - that box's captions are 继续/结束 and pressing IDOK there
    ' would CONTINUE the program instead of dismissing it.
    If WoMgIdFallbackAllowed(d, btns) Then
        Dim pass As Long, wantId As Long
        For pass = 1 To 3
            If pass = 1 Then
                wantId = 2
            ElseIf pass = 2 Then
                wantId = 1
            Else
                wantId = 3
            End If
            For j = 1 To btns.Count
                Set b = btns.Item(j)
                If CBool(b("enabled")) Then
                    If CLng(b("id")) = wantId Then
                        outText = CStr(b("text"))
                        outCancel = (wantId = 2)
                        outWhy = "standard_control_id_fallback"
                        outId = wantId
                        WoMgPickButton = CLng(b("hwnd"))
                        Exit Function
                    End If
                End If
            Next
        Next
    End If
    outWhy = "no_safe_button"
    Err.Clear
End Function

' TRUE only when standard-ID fallback cannot CONTINUE or END a program: the body must
' not be a runtime-error box, and no caption may be End/Continue in either language.
Private Function WoMgButtonsOf(ByVal d As Object) As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    Dim btns As Object
    Set btns = btnsOf(d)
    If btns Is Nothing Then Set WoMgButtonsOf = out : Err.Clear : Exit Function
    Dim j As Long, b As Object, e As Object
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        Set e = CreateObject("Scripting.Dictionary")
        e("text") = CStr(b("text"))
        e("id") = CLng(b("id"))
        e("hwnd") = CLng(b("hwnd"))
        e("enabled") = CBool(b("enabled"))
        out.Add e
    Next
    Err.Clear
    Set WoMgButtonsOf = out
End Function

Private Function WoMgIdFallbackAllowed(ByVal d As Object, ByVal btns As Object) As Boolean
    WoMgIdFallbackAllowed = True
    On Error Resume Next
    If WoDlgKindOf(d) = "runtime_error" Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
    On Error Resume Next
    Dim body As String
    body = UCase$(WoMgBodyOf(d))
    If InStr(1, body, "RUN-TIME ERROR") > 0 Or InStr(1, body, "RUNTIME ERROR") > 0 Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
    If InStr(1, body, ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
    If InStr(1, body, ChrW$(&H5B9E) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
    Dim j As Long, b As Object, t As String, n As String
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        t = CStr(b("text"))
        n = UCase$(Trim$(Replace(t, "&", "")))
        If n = "END" Or n = "CONTINUE" Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
        If InStr(1, t, ChrW$(&H7ED3) & ChrW$(&H675F)) > 0 Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
        If InStr(1, t, ChrW$(&H7EE7) & ChrW$(&H7EED)) > 0 Then WoMgIdFallbackAllowed = False : Err.Clear : Exit Function
    Next
    Err.Clear
End Function

' FIX-7 P0: DEFAULT POLICY TABLE. Category first, then the action this guard may take.
' runtime_error      -> click DEBUG (only when the armed tool is NOT runtime-aware, because
'                       those tools have their own READ->DEBUG capture path and must win);
' compile/reference/component_missing/breakpoint_refused -> click OK;
' save_confirm       -> Cancel (or No only when the arm explicitly allows YES/NO IDs);
' host_modal_risk=known -> HOLD (existing allow_host_modal rule unchanged);
' unknown / self-drawn -> HOLD (awaiting_user) - never a silent click.
Private Function WoMgClassify(ByVal d As Object, ByRef category As String, ByRef action As String, Optional ByVal preExisting As Boolean = False) As Boolean
    On Error Resume Next
    category = "unknown"
    ' FIX-11: an UNRECOGNISED box must not strand the call. Press the safe CANCEL/Close
    ' (IDCANCEL=2) when one exists; the tick releases the gate when there is nothing safe.
    action = "click_cancel"
    Dim bdy As String
    bdy = UCase$(WoMgBodyOf(d))
    Dim ttl As String
    ttl = UCase$(CStr(d("title")))
    Dim risk As String, rWhy As String, rCls As String
    WoHostModalRiskOfDialog d, risk, rWhy, rCls
    If risk = "known" Then
        category = "host_modal_risk_known"
        ' FIX-11: cancel-like is always safe here; the risk branch below still refuses a
        ' non-cancel button, and the tick falls back to a gate release.
        action = "click_cancel"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    ' FIX-10: the DEBUGGED PROGRAM's own MsgBox. Conservative three-condition test:
    '  (1) the IDE is RUNNING (GuardIdeState = run, or EbMode=1 outside break/design);
    '  (2) the box is a #32770 whose owner is DISABLED (a modal MsgBox blocks its owner);
    '  (3) the owner is NOT the IDE main window - a VB6 runtime form class (ThunderRT6...)
    '      or an owner caption that differs from the IDE main caption.
    ' Any condition that cannot be established means the box stays "unknown". The policy is
    ' still HOLD: answering the program's business question on the user's behalf is unsafe.
    If WoMgRunState() Then
        If CBool(d("ownerDisabled")) Then
            Dim pgOwn As Long
            Dim pgOwnCls As String
            Dim pgOwnTitle As String
            pgOwn = 0
            pgOwnCls = ""
            pgOwnTitle = ""
            If d.Exists("owner") Then pgOwn = CLng(d("owner"))
            If d.Exists("ownerCls") Then pgOwnCls = CStr(d("ownerCls"))
            If d.Exists("ownerTitle") Then pgOwnTitle = CStr(d("ownerTitle"))
            If pgOwn <> 0 Then
                Dim pgIsProgram As Boolean
                pgIsProgram = False
                If LCase$(Left$(pgOwnCls, 9)) = "thunderrt" Then
                    pgIsProgram = True
                ElseIf Len(pgOwnTitle) > 0 Then
                    If StrComp(pgOwnTitle, WoMgMainTitle(), vbTextCompare) <> 0 Then pgIsProgram = True
                End If
                If pgIsProgram Then
                    category = "program_msgbox"
                    action = "hold"
                    WoMgClassify = True
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
    End If
    ' FIX-18: STRUCTURAL run-mode discriminator. The IDE's OWN runtime-error box always has a
    ' Debug button (the measured trio is Debug/End/Help); the DEBUGGED PROGRAM's MsgBox has no
    ' such button - and its owner CAN be the IDE main window (Sub Main has no form), so the
    ' owner-based three-condition test above cannot see it. A run-mode box with NO Debug
    ' button is therefore ALWAYS the program's own box -> HOLD (never answer it for the user).
    If WoMgRunState() Then
        If Not WoMgHasDebugButton(d) Then
            category = "program_msgbox"
            action = "hold"
            d("sessionMsgBox") = True
            d("programMsgBoxEvidence") = "no_debug_button_in_run_mode"
            WoMgClassify = True
            Err.Clear
            Exit Function
        End If
    End If
    If InStr(1, bdy, "RUN-TIME ERROR") > 0 Or InStr(1, bdy, "RUNTIME ERROR") > 0 _
       Or InStr(1, bdy, ChrW$(&H8FD0) & ChrW$(&H884C) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 _
       Or InStr(1, bdy, ChrW$(&H5B9E) & ChrW$(&H65F6) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then
        category = "runtime_error"
        ' FIX-9: a runtime-error box that was ALREADY up when we arrived is still pressed
        ' on DEBUG even when the current tool is runtime-aware: otherwise every subsequent
        ' call keeps being blocked by a box nobody dismisses. save_confirm is unaffected.
        If m_mgAllowRuntimeDebug Or preExisting Then action = "click_debug" Else action = "hold"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, bdy, "COMPILE ERROR") > 0 Or InStr(1, bdy, ChrW$(&H7F16) & ChrW$(&H8BD1) & ChrW$(&H9519) & ChrW$(&H8BEF)) > 0 Then
        category = "compile"
        action = "click_ok"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, bdy, "REFERENCE") > 0 Or InStr(1, bdy, ChrW$(&H5F15) & ChrW$(&H7528)) > 0 Then
        category = "reference"
        action = "click_ok"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, bdy, "NOT FOUND") > 0 Or InStr(1, bdy, ChrW$(&H672A) & ChrW$(&H627E) & ChrW$(&H5230)) > 0 _
       Or InStr(1, bdy, ChrW$(&H7EC4) & ChrW$(&H4EF6)) > 0 Then
        category = "component_missing"
        action = "click_ok"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, bdy, ChrW$(&H65AD) & ChrW$(&H70B9)) > 0 Or InStr(1, bdy, "BREAKPOINT") > 0 Then
        category = "breakpoint_refused"
        action = "click_ok"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    If InStr(1, ttl, "SAVE") > 0 Or InStr(1, ttl, ChrW$(&H4FDD) & ChrW$(&H5B58)) > 0 _
       Or InStr(1, bdy, ChrW$(&H53E6) & ChrW$(&H5B58)) > 0 Then
        category = "save_confirm"
        If m_mgAllowIdYesNo Then action = "click_no" Else action = "click_cancel"
        WoMgClassify = True
        Err.Clear
        Exit Function
    End If
    WoMgClassify = True
    Err.Clear
End Function

Private Function WoMgButtonByCaption(ByVal d As Object, ByVal normWant As String, ByVal cp1 As Long, ByVal cp2 As Long, ByRef outText As String, ByRef outId As Long) As Long
    On Error Resume Next
    outText = ""
    outId = 0
    Dim btns As Object
    Set btns = btnsOf(d)
    If btns Is Nothing Then Err.Clear : Exit Function
    Dim j As Long, b As Object, t As String, n As String
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        If CBool(b("enabled")) Then
            t = CStr(b("text"))
            n = UCase$(Trim$(Replace(t, "&", "")))
            If n = normWant Or (cp1 <> 0 And InStr(1, t, ChrW$(cp1) & ChrW$(cp2)) > 0) Then
                outText = t
                outId = CLng(b("id"))
                WoMgButtonByCaption = CLng(b("hwnd"))
                Err.Clear
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

Private Function WoMgPickForAction(ByVal d As Object, ByVal action As String, ByRef outText As String, ByRef outWhy As String, ByRef outCancel As Boolean, ByRef outId As Long) As Long
    On Error Resume Next
    outText = ""
    outWhy = action
    outCancel = (action = "click_cancel")
    outId = 0
    Dim btns As Object
    Set btns = btnsOf(d)
    If btns Is Nothing Then Err.Clear : Exit Function
    Dim j As Long, b As Object, t As String, n As String
    If action = "click_debug" Then
        ' ONLY the DEBUG caption (name or the localized code points). No generic fallback:
        ' if the box has no DEBUG button the policy must HOLD/report, never press OK.
        WoMgPickForAction = WoMgButtonByCaption(d, "DEBUG", &H8C03, &H8BD5, outText, outId)
        If WoMgPickForAction = 0 Then outWhy = "no_debug_button"
        Err.Clear
        Exit Function
    End If
    For j = 1 To btns.Count
        Set b = btns.Item(j)
        If CBool(b("enabled")) Then
            t = CStr(b("text"))
            n = UCase$(Trim$(Replace(t, "&", "")))
            Dim idNow As Long
            idNow = CLng(b("id"))
            Dim ok As Boolean
            ok = False
            If action = "click_ok" Then
                If idNow = 1 Or n = "OK" Or n = "OKAY" Or n = ChrW$(&H786E) & ChrW$(&H5B9A) Then ok = True
            ElseIf action = "click_cancel" Then
                If idNow = 2 Or WoButtonCancelLike(n) Or n = ChrW$(&H53D6) & ChrW$(&H6D88) Then ok = True
            ElseIf action = "click_no" Then
                If m_mgAllowIdYesNo Then
                    If idNow = 7 Or n = "NO" Or n = ChrW$(&H5426) Then ok = True
                End If
            ElseIf action = "click_debug" Then
                ' unreachable for click_debug (handled above); kept out of the safe-caption
                ' fallback on purpose
            End If
            If ok Then
                outText = t
                outId = idNow
                WoMgPickForAction = CLng(b("hwnd"))
                Err.Clear
                Exit Function
            End If
        End If
    Next
    Err.Clear
End Function

' FIX-10: pick the button a HUMAN may safely press from this ledger entry: standard
' IDCANCEL=2 first, then IDOK=1, then IDABORT=3, then a safe caption. 0 = none found (the
' caller must then read the true hwnd via read_active_dialogs).
Private Function WoRecoverySafeHwnd(ByVal buttons As Object) As Long
    WoRecoverySafeHwnd = 0
    On Error Resume Next
    If buttons Is Nothing Then Err.Clear : Exit Function
    Dim pass As Long
    Dim want As Long
    Dim i As Long
    Dim b As Object
    For pass = 1 To 3
        If pass = 1 Then
            want = 2
        ElseIf pass = 2 Then
            want = 1
        Else
            want = 3
        End If
        For i = 1 To buttons.Count
            Set b = buttons.Item(i)
            If CLng(b("id")) = want Then
                WoRecoverySafeHwnd = CLng(b("hwnd"))
                Err.Clear
                Exit Function
            End If
        Next
    Next
    Dim t As String
    Dim n As String
    For i = 1 To buttons.Count
        Set b = buttons.Item(i)
        t = CStr(b("text"))
        n = ""
        If WoButtonAllow(t, n) Then
            WoRecoverySafeHwnd = CLng(b("hwnd"))
            Err.Clear
            Exit Function
        End If
    Next
    Err.Clear
End Function

' FIX-10: the executable self-recovery block. Only produced for a HELD / UNRESOLVED dialog;
' it is additive next to the prose howToRecover. hwnd is a MEASURED value from the ledger's
' buttons[]; when none can be chosen the step says where to read it instead.
Public Function WoRecoveryForButtons(ByVal buttons As Object, ByVal category As String) As Object
    Set WoRecoveryForButtons = Nothing
    On Error Resume Next
    Dim r As Object
    Set r = CreateObject("Scripting.Dictionary")
    Dim steps As Object
    Set steps = New Collection
    Dim s1 As Object
    Set s1 = CreateObject("Scripting.Dictionary")
    s1("tool") = "read_active_dialogs"
    Set s1("args") = CreateObject("Scripting.Dictionary")
    steps.Add s1
    Dim s2 As Object
    Set s2 = CreateObject("Scripting.Dictionary")
    s2("tool") = "send_window_message"
    Dim a2 As Object
    Set a2 = CreateObject("Scripting.Dictionary")
    a2("action") = "click"
    Dim safeH As Long
    safeH = WoRecoverySafeHwnd(buttons)
    If safeH <> 0 Then
        a2("hwnd") = safeH
        s2("hwndFrom") = "ledger"
    Else
        s2("hwndFrom") = "read_active_dialogs"
    End If
    s2("note") = "only captions/IDs on the safe list; NEVER End/Continue"
    Set s2("args") = a2
    steps.Add s2
    Dim s3 As Object
    Set s3 = CreateObject("Scripting.Dictionary")
    s3("tool") = "ide_state"
    Dim a3 As Object
    Set a3 = CreateObject("Scripting.Dictionary")
    a3("release_inflight") = True
    a3("release_min_held_ms") = 5000
    Set s3("args") = a3
    steps.Add s3
    Set r("steps") = steps
    Dim unsafeList As Object
    Set unsafeList = New Collection
    unsafeList.Add ChrW$(&H7ED3) & ChrW$(&H675F) & "(&E)"
    unsafeList.Add ChrW$(&H7EE7) & ChrW$(&H7EED) & "(&C)"
    Set r("guaranteedUnsafe") = unsafeList
    r("reason") = "dialog could not be resolved by the guard (" & category & ")"
    Set WoRecoveryForButtons = r
    Err.Clear
End Function

Public Function WoRecoveryForDialog(ByVal d As Object) As Object
    Set WoRecoveryForDialog = Nothing
    On Error Resume Next
    If d Is Nothing Then Err.Clear : Exit Function
    Dim bs As Object
    Set bs = btnsOf(d)
    Set WoRecoveryForDialog = WoRecoveryForButtons(bs, "unknown")
    Err.Clear
End Function

Public Function WoModalRecovery() As Object
    Set WoModalRecovery = Nothing
    On Error Resume Next
    If Not (m_mgAwaiting Or m_mgUnresolvable) Then Err.Clear : Exit Function
    Dim btns As Object
    Set btns = m_mgButtons
    If btns Is Nothing Then
        If Not m_mgDialogs Is Nothing Then
            If m_mgDialogs.Count > 0 Then
                Dim lastD As Object
                Set lastD = m_mgDialogs(m_mgDialogs.Count)
                If lastD.Exists("buttons") Then Set btns = lastD("buttons")
            End If
        End If
    End If
    Set WoModalRecovery = WoRecoveryForButtons(btns, m_mgCategory)
    Err.Clear
End Function

' FIX-12: a short PUMPING wait (DoEvents lets the host modal loop - and therefore our own
' tick - run while we verify a click).
Private Sub WoMgPause(ByVal ms As Long)
    Dim t0 As Long
    t0 = WoGetTickCount()
    Do While WoGetTickCount() - t0 < ms
        m_pfPumpT0 = WoProfNow()
        DoEvents
        WoPumpAdd WoProfNow() - m_pfPumpT0
        WoKernelSleep 10
    Loop
End Sub

' FIX-21: is there a pre-existing owner-DISABLED #32770 still on screen? This is a PURE
' Win32 enumeration/read (no VBIDE/CommandBars call), so it is SAFE to call while a modal box
' owns the thread. A HOLD-policy box (the program's own MsgBox) is never dismissed by the
' guard, so it stays here and must block later IDE calls instead of letting them time out.
' FIX-21b: MESSAGE-FREE scan. EnumWindows + GetWindowThreadProcessId + GetClassName +
' IsWindowVisible + GetWindow/IsWindowEnabled only - NOTHING here sends a window message
' (no GetWindowText, no WM_GETTEXT, no Click). That matters because a synchronous
' GetWindowText can block forever while the IDE thread is busy/suspended.
Public Function WoModalFastPresent(ByRef dialogCountOut As Long) As Boolean
    WoModalFastPresent = False
    dialogCountOut = 0
    On Error Resume Next
    WoScanTops
    Dim n As Long
    n = m_woCount
    If n > WO_MAX Then n = WO_MAX
    If n <= 0 Then Err.Clear : Exit Function
    Dim pid As Long
    pid = WoGetCurrentProcessId()
    Dim i As Long
    For i = 0 To n - 1
        Dim h As Long
        h = m_woHwnd(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                If LCase$(WoCls(h)) = "#32770" Then
                    If WoIsWindowVisible(h) <> 0 Then
                        Dim own As Long
                        own = WoGetWindow(h, GW_OWNER)
                        If own <> 0 Then
                            If WoIsWindowEnabled(own) = 0 Then
                                dialogCountOut = dialogCountOut + 1
                                WoModalFastPresent = True
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function

Public Function WoModalBlockingBox() As Object
    Set WoModalBlockingBox = Nothing
    On Error Resume Next
    Dim cnt As Long
    cnt = 0
    If Not WoModalFastPresent(cnt) Then Err.Clear : Exit Function
    ' Locate the first owner-disabled box WITHOUT reading any title/body.
    WoScanTops
    Dim n As Long
    n = m_woCount
    If n > WO_MAX Then n = WO_MAX
    Dim pid As Long
    pid = WoGetCurrentProcessId()
    Dim i As Long
    For i = 0 To n - 1
        Dim h As Long
        h = m_woHwnd(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                If LCase$(WoCls(h)) = "#32770" Then
                    If WoIsWindowVisible(h) <> 0 Then
                        Dim own As Long
                        own = WoGetWindow(h, GW_OWNER)
                        If own <> 0 Then
                            If WoIsWindowEnabled(own) = 0 Then
                                Dim out As Object
                                Set out = CreateObject("Scripting.Dictionary")
                                out("hwnd") = h
                                out("cls") = WoCls(h)
                                out("title") = ""
                                out("titleDeferred") = True
                                out("classification") = "preexisting_modal_unknown"
                                out("policyAction") = "hold"
                                out("held") = True
                                out("heldUnknown") = True
                                out("dialogCount") = cnt
                                Set WoModalBlockingBox = out
                                Err.Clear
                                Exit Function
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
End Function

' FIX-14: a modal box left on screen with NO in-flight call (e.g. a compile-error box the
' caller never dismissed) is handled at the ENTRY of the next tool call. Strategy comes from
' the same classifier: compile/reference -> OK; save_confirm -> Cancel; runtime_error -> DEBUG;
' unknown/host-modal -> Cancel. End/Continue are ALWAYS refused.
Public Function WoModalClearPreexisting() As Object
    Set WoModalClearPreexisting = Nothing
    On Error Resume Next
    Dim ds As Object
    Set ds = WoDialogsJson("#32770", 0, True)
    If ds Is Nothing Then Err.Clear : Exit Function
    Dim i As Long
    For i = 1 To ds.Count
        Dim d As Object
        Set d = ds.Item(i)
        If CBool(d("ownerDisabled")) Then
            Dim cat As String, act As String
            cat = "unknown"
            act = "hold"
            WoMgClassify d, cat, act
            Dim pickAct As String
            ' FIX-15: the classifier is the SINGLE policy source. This entry handler must not
            ' invent its own default (that turned program_msgbox's HOLD into a Cancel click).
            pickAct = act
            If cat = "runtime_error" Then pickAct = "click_debug"
            If cat = "compile" Or cat = "reference" Or cat = "component_missing" Or cat = "breakpoint_refused" Then pickAct = "click_ok"
            If cat = "save_confirm" Then pickAct = "click_cancel"
            If cat = "program_msgbox" Then pickAct = "hold"
            ' never press End/Continue, whatever the class says
            Dim btns As Object
            Set btns = btnsOf(d)
            If Not btns Is Nothing Then
                Dim j As Long, b As Object, t As String, nn As String
                Dim forbid As Boolean
                forbid = False
                For j = 1 To btns.Count
                    Set b = btns.Item(j)
                    t = CStr(b("text"))
                    nn = UCase$(Trim$(Replace(t, "&", "")))
                    If nn = "END" Or nn = "CONTINUE" Then forbid = True
                    If InStr(1, t, ChrW$(&H7ED3) & ChrW$(&H675F)) > 0 Then forbid = True
                    If InStr(1, t, ChrW$(&H7EE7) & ChrW$(&H7EED)) > 0 Then forbid = True
                Next
                If forbid Then pickAct = "hold"
            End If
            Dim pH As Long, pId As Long, pT As String, pw As String, pc As Boolean
            pT = ""
            pw = ""
            pc = False
            pId = 0
            Dim mth As String, res As String
            mth = "none"
            res = "refused"
            Dim heldNow As Boolean
            heldNow = False
            If pickAct = "hold" Or pickAct = "click_no" Then heldNow = True
            If heldNow Then
                ' policy HOLD: send NOTHING at all, and say so (never ok=true).
                mth = "none"
                res = "held_policy"
            Else
                If pickAct = "click_cancel" Then
                    pH = WoMgPickForAction(d, "click_cancel", pT, pw, pc, pId)
                ElseIf pickAct = "click_ok" Then
                    pH = WoMgPickForAction(d, "click_ok", pT, pw, pc, pId)
                ElseIf pickAct = "click_debug" Then
                    pH = WoMgPickForAction(d, "click_debug", pT, pw, pc, pId)
                End If
                If pH <> 0 Then
                    Call WoMgPressVerify(d, pH, pId, pT, mth, res)
                End If
            End If
            Dim out As Object
            Set out = CreateObject("Scripting.Dictionary")
            out("hwnd") = d("hwnd")
            out("title") = d("title")
            out("classification") = cat
            out("plannedAction") = pickAct
            out("clickedId") = pId
            out("clickedCaption") = pT
            out("method") = mth
            out("result") = res
            out("closedNow") = (WoIsWindow(CLng(d("hwnd"))) = 0)
            out("preExisting") = True
            If heldNow Then
                WoMgRecordDialog d, cat, "preexisting_held", 0, "", "hold", "held_policy"
                Call WoMgRecordTick("preexisting_hold", CLng(d("hwnd")), 0, "", "none", "held_policy")
            Else
                WoMgRecordDialog d, cat, "preexisting_handled", pId, pT, pickAct, res
                Call WoMgRecordTick("preexisting_handled", CLng(d("hwnd")), pId, pT, mth, res)
            End If
            Set WoModalClearPreexisting = out
            Err.Clear
            Exit Function
        End If
    Next
    Err.Clear
End Function

' FIX-13: process-level observability that survives a wedged call (the call's own receipt
' never reaches the caller, so ide_state / mcp_info / the fast path must expose this).
Public Function WoModalGuardResident() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("armed") = m_mgArmed
    d("tickCount") = m_mgTickCount
    d("refusedTicks") = m_mgRefused
    d("modalCloseSent") = m_mgCloseSent
    d("firstSeenAtMs") = m_mgFirstSeenAt
    d("deadlineAtMs") = IIf(m_mgFirstSeenAt = 0, 0, m_mgFirstSeenAt + MODAL_WEDGE_DEADLINE_MS)
    If m_mgLastTickAction Is Nothing Then
        d("lastTickAction") = "none"
    Else
        Set d("lastTickAction") = m_mgLastTickAction
    End If
    Set WoModalGuardResident = d
End Function

' FIX-13: ONE safe-close implementation - this calls the SAME close_active_dialogs code that
' the live run proved can dismiss the Add-Form box (caption -> ID fallback -> BM_CLICK ->
' WM_COMMAND -> recheck). Runtime-error/End/Continue boxes are refused before the call.
Public Function WoCloseDialogSafe(ByVal hwnd As Long) As Object
    Set WoCloseDialogSafe = Nothing
    On Error Resume Next
    Dim ds As Object
    Set ds = WoDialogsJson("#32770", hwnd, True)
    If ds Is Nothing Then Err.Clear : Exit Function
    If ds.Count < 1 Then Err.Clear : Exit Function
    Dim d As Object
    Set d = ds.Item(1)
    Dim btns As Object
    Set btns = btnsOf(d)
    If Not btns Is Nothing Then
        Dim i As Long, b As Object, t As String, nn As String
        For i = 1 To btns.Count
            Set b = btns.Item(i)
            t = CStr(b("text"))
            nn = UCase$(Trim$(Replace(t, "&", "")))
            If nn = "END" Or nn = "CONTINUE" Then Err.Clear : Exit Function
            If InStr(1, t, ChrW$(&H7ED3) & ChrW$(&H675F)) > 0 Then Err.Clear : Exit Function
            If InStr(1, t, ChrW$(&H7EE7) & ChrW$(&H7EED)) > 0 Then Err.Clear : Exit Function
        Next
    End If
    Dim raw As String
    raw = CloseActiveDialogsJson(hwnd, "", False, False)
    Dim v As Variant
    Dim out As Object
    Set out = CreateObject("Scripting.Dictionary")
    out("raw") = Left$(raw, 4000)
    out("clickedId") = 0
    out("clickedCaption") = ""
    out("method") = ""
    out("result") = "still_present"
    If JsonParse(raw, v) Then
        If IsObject(v) Then
            Dim rj As Object
            Set rj = v
            If rj.Exists("clickedCount") Then
                If CLng(rj("clickedCount")) > 0 Then
                    Dim clk As Object
                    Set clk = rj("clicked")
                    If clk.Count > 0 Then
                        Dim c1 As Object
                        Set c1 = clk.Item(1)
                        If c1.Exists("clickedId") Then out("clickedId") = c1("clickedId")
                        If c1.Exists("text") Then out("clickedCaption") = c1("text")
                        If c1.Exists("methodUsed") Then out("method") = c1("methodUsed")
                    End If
                End If
            End If
            If rj.Exists("dialogsAfter") Then
                If CLng(rj("dialogsAfter")) = 0 Then out("result") = "ok"
            End If
        End If
    End If
    out("stillPresent") = (WoIsWindow(hwnd) <> 0)
    Set WoCloseDialogSafe = out
    Err.Clear
End Function

' FIX-12: record EVERY attempt so a caller can prove what the tick did (or that it never ran).
Private Sub WoMgRecordTick(ByVal action As String, ByVal hwnd As Long, ByVal buttonId As Long, ByVal caption As String, ByVal method As String, ByVal result As String)
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("atMs") = IIf(m_mgAt = 0, 0, WoGetTickCount() - m_mgAt)
    d("action") = action
    d("hwnd") = hwnd
    d("buttonId") = buttonId
    d("buttonCaption") = caption
    d("method") = method
    d("result") = result
    d("sawBoxCount") = m_mgSeenCount
    d("ticks") = m_mgTickCount
    Set m_mgLastTickAction = d
    Err.Clear
End Sub

' FIX-12: the ONE click method (same as close_active_dialogs): BM_CLICK, re-check after ~300 ms,
' then WM_COMMAND/BN_CLICKED on the DIALOG, then ONE WM_CLOSE, re-checking at every step.
Private Function WoMgPressVerify(ByVal d As Object, ByVal pickHwnd As Long, ByVal pickId As Long, ByVal pickText As String, ByRef methodOut As String, ByRef resultOut As String) As Boolean
    WoMgPressVerify = False
    On Error Resume Next
    Dim dlg As Long
    dlg = 0
    If d Is Nothing Then Err.Clear : Exit Function
    dlg = CLng(d("hwnd"))
    methodOut = "BM_CLICK"
    resultOut = "ok"
    WoSendMessageW pickHwnd, BM_CLICK, 0, 0
    Call WoMgPause(300)
    If WoIsWindow(dlg) = 0 Then
        WoMgPressVerify = True
        Err.Clear
        Exit Function
    End If
    WoSendMessageW dlg, WM_COMMAND, pickId, pickHwnd
    methodOut = "BM_CLICK+WM_COMMAND"
    Call WoMgPause(300)
    If WoIsWindow(dlg) = 0 Then
        WoMgPressVerify = True
        Err.Clear
        Exit Function
    End If
    If Not m_mgCloseSent Then
        WoSendMessageW dlg, WM_CLOSE, 0, 0
        m_mgCloseSent = True
        methodOut = methodOut & "+WM_CLOSE"
        Call WoMgPause(300)
        If WoIsWindow(dlg) = 0 Then
            WoMgPressVerify = True
            Err.Clear
            Exit Function
        End If
    End If
    resultOut = "still_present"
    Err.Clear
End Function

' FIX-12: cancel a modal box that was ALREADY open before an interactive tool ran.
Public Function WoModalCancelPriorDialog() As Object
    Set WoModalCancelPriorDialog = Nothing
    On Error Resume Next
    Dim ds As Object
    Set ds = WoDialogsJson("#32770", 0, True)
    If ds Is Nothing Then Err.Clear : Exit Function
    Dim i As Long
    For i = 1 To ds.Count
        Dim d As Object
        Set d = ds.Item(i)
        If CBool(d("ownerDisabled")) Then
            Dim cat As String, act As String
            cat = "unknown"
            act = "hold"
            WoMgClassify d, cat, act
            Dim pH As Long, pId As Long, pT As String, pw As String, pc As Boolean
            pT = ""
            pw = ""
            pc = False
            pId = 0
            pH = WoMgPickForAction(d, "click_cancel", pT, pw, pc, pId)
            If pH = 0 Then pH = WoMgPickForAction(d, "click_ok", pT, pw, pc, pId)
            Dim mth As String, res As String
            mth = "none"
            res = "refused"
            If pH <> 0 Then
                Call WoMgPressVerify(d, pH, pId, pT, mth, res)
                m_mgClickedHwnd = pH
                m_mgClickedText = pT
                m_mgClickedId = pId
                m_mgClicked = (res = "ok")
            End If
            Dim out As Object
            Set out = CreateObject("Scripting.Dictionary")
            out("hwnd") = d("hwnd")
            out("title") = d("title")
            out("classification") = cat
            out("clickedId") = pId
            out("clickedCaption") = pT
            out("method") = mth
            out("result") = res
            out("closedNow") = (WoIsWindow(CLng(d("hwnd"))) = 0)
            out("buttons") = d("buttons")
            Call WoMgRecordTick("cancel_prior", CLng(d("hwnd")), pId, pT, mth, res)
            Set WoModalCancelPriorDialog = out
            Err.Clear
            Exit Function
        End If
    Next
    Err.Clear
End Function

' FIX-11: hard-deadline helpers.
Private Function WoMgDeadlineExceeded() As Boolean
    WoMgDeadlineExceeded = False
    If m_mgFirstSeenAt = 0 Then Exit Function
    If (WoGetTickCount() - m_mgFirstSeenAt) >= MODAL_WEDGE_DEADLINE_MS Then WoMgDeadlineExceeded = True
End Function

Private Sub WoMgForceUnstick(ByVal d As Object)
    On Error Resume Next
    Dim safeRes As Object
    Set safeRes = WoCloseDialogSafe(CLng(d("hwnd")))
    If Not safeRes Is Nothing Then
        Dim sid As Long, scap As String, smth As String, sres As String
        sid = CLng(safeRes("clickedId"))
        scap = CStr(safeRes("clickedCaption"))
        smth = CStr(safeRes("method"))
        sres = CStr(safeRes("result"))
        m_mgClickedId = sid
        m_mgClickedText = scap
        m_mgMethod = smth & " (hard deadline, shared closer)"
        Call WoMgRecordTick("click_safe", CLng(d("hwnd")), sid, scap, smth, sres)
        WoMgRecordDialog d, m_mgCategory, "click_deadline", sid, scap, "click_cancel", sres
        If sres = "ok" Then
            m_mgClickedHwnd = 0
            m_mgClicked = True
            m_mgArmed = False
            Err.Clear
            Exit Sub
        End If
    End If
    If Not m_mgCloseSent Then
        ' one-shot, parameterless, this window only
        WoSendMessageW CLng(d("hwnd")), WM_CLOSE, 0, 0
        m_mgCloseSent = True
        Call WoMgPause(300)
        m_mgMethod = "WM_CLOSE (hard deadline, one shot)"
        m_mgAwaiting = True
        WoMgRecordDialog d, m_mgCategory, "close_deadline", 0, "", "close", IIf(WoIsWindow(CLng(d("hwnd"))) = 0, "ok", "still_present")
        Call WoMgRecordTick("close", CLng(d("hwnd")), 0, "", "WM_CLOSE", IIf(WoIsWindow(CLng(d("hwnd"))) = 0, "ok", "still_present"))
        Err.Clear
        Exit Sub
    End If
    Call WoMgWedgeAbort("hard_deadline: " & m_mgCategory)
    Call WoMgRecordTick("release", CLng(d("hwnd")), 0, "", "gate", IIf(m_mgGateReleased, "ok", "refused"))
    m_mgMethod = "gate release (hard deadline)"
    Err.Clear
End Sub

' FIX-11: release the single-flight gate FROM THE TICK. This is the only code that runs
' while the host's modal loop owns the thread. The wedge is recorded so the (eventually
' returned) receipt and every later receipt can report it as a structured fact.
Private Function WoMgWedgeAbort(ByVal why As String) As Boolean
    WoMgWedgeAbort = False
    On Error Resume Next
    m_mgWedgeAborted = True
    m_mgWedgeWhy = why
    m_mgWedgeAt = IIf(m_mgAt = 0, 0, WoGetTickCount() - m_mgAt)
    m_mgAwaiting = True
    m_mgGateReleased = False
    If Not m_whSink Is Nothing Then
        m_mgGateReleased = CBool(m_whSink.GuardWedgeAbort(m_mgCategory, why))
    End If
    WoMgWedgeAbort = m_mgGateReleased
    Err.Clear
End Function

Public Function WoModalGuardWedgeAborted() As Boolean
    WoModalGuardWedgeAborted = m_mgWedgeAborted
End Function

Public Function WoModalGuardCancelClicked() As Boolean
    WoModalGuardCancelClicked = (m_mgClicked And m_mgClickedId = 2)
End Function

Public Function WoModalGuardLastClickedId() As Long
    WoModalGuardLastClickedId = m_mgClickedId
End Function

Public Function WoModalWedgeJson() As Object
    Set WoModalWedgeJson = Nothing
    On Error Resume Next
    If Not m_mgWedgeAborted Then Err.Clear : Exit Function
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("atMs") = m_mgWedgeAt
    d("category") = m_mgCategory
    d("why") = m_mgWedgeWhy
    d("gateReleased") = m_mgGateReleased
    d("closeSent") = m_mgCloseSent
    d("releasedBy") = "modal_guard_tick"
    Dim bs As Object
    Set bs = New Collection
    Dim i As Long
    Dim b As Object
    If Not m_mgButtons Is Nothing Then
        For i = 1 To m_mgButtons.Count
            Set b = m_mgButtons(i)
            bs.Add CStr(b("text"))
        Next
    End If
    Set d("buttons") = bs
    Set WoModalWedgeJson = d
    Err.Clear
End Function

' FIX-11: a MINIMAL ledger snapshot for receipts produced while the guard is armed but
' nothing was recorded yet (a nested read served at a pump point). The boxes are real
' (enumerated now) and say preExisting=true - never a fabricated cause.
Public Function WoModalSnapshot() As Object
    Dim out As Object
    Set out = New Collection
    On Error Resume Next
    ' FIX-12: do NOT require the guard to be armed - during a wedge it may already have
    ' disarmed, and that is exactly when the caller needs the live box in dialogs[].
    Dim ds As Object
    Set ds = WoDialogsJson("", 0, True)
    Dim i As Long
    For i = 1 To ds.Count
        Dim d As Object
        Dim e As Object
        Set d = ds.Item(i)
        Set e = CreateObject("Scripting.Dictionary")
        e("hwnd") = d("hwnd")
        e("cls") = d("cls")
        e("title") = d("title")
        e("bodyHead") = Left$(CStr(d("title")), 200)
        e("ownerDisabled") = d("ownerDisabled")
        e("buttons") = d("buttons")
        e("classification") = "unknown"
        e("action") = "held"
        e("clickedId") = 0
        e("latencyMs") = IIf(m_mgAt = 0, 0, WoGetTickCount() - m_mgAt)
        e("source") = "modal_guard_snapshot"
        e("preExisting") = True
        out.Add e
    Next
    Err.Clear
    Set WoModalSnapshot = out
End Function

' FIX-10: the PROGRAM's own MsgBox check needs the IDE state and the IDE main-window
' caption. Both live in the late-bound VBIDE sink (m_whSink) that McpServer wires in; when
' it is not available the classification conservatively says "not running".
Private Function WoMgRunState() As Boolean
    WoMgRunState = False
    On Error Resume Next
    If m_whSink Is Nothing Then Err.Clear : Exit Function
    Dim st As String
    st = ""
    st = LCase$(CStr(m_whSink.GuardIdeState()))
    If Len(st) = 0 Then Err.Clear : Exit Function
    If st = "run" Then
        WoMgRunState = True
    ElseIf Vba6Probe.IdeEbMode() = 1 Then
        If st <> "break" Then
            If st <> "design" Then WoMgRunState = True
        End If
    End If
    Err.Clear
End Function

Private Function WoMgMainTitle() As String
    WoMgMainTitle = ""
    On Error Resume Next
    If m_whSink Is Nothing Then Err.Clear : Exit Function
    WoMgMainTitle = CStr(m_whSink.GuardMainTitle())
    Err.Clear
End Function

' FIX-9: was this dialog already on screen when the guard was armed?
Private Function WoMgIsPre(ByVal h As Long) As Boolean
    WoMgIsPre = False
    On Error Resume Next
    If m_mgPre Is Nothing Then Err.Clear : Exit Function
    Dim i As Long
    For i = 1 To m_mgPre.Count
        If CLng(m_mgPre(i)) = h Then
            WoMgIsPre = True
            Exit Function
        End If
    Next
    Err.Clear
End Function

' FIX-18: does this box carry the IDE runtime-error Debug button? Normalised caption
' compare (Debug / 调试), never a body-text guess.
Private Function WoMgHasDebugButton(ByVal d As Object) As Boolean
    WoMgHasDebugButton = False
    On Error Resume Next
    Dim bs As Object
    Set bs = btnsOf(d)
    If bs Is Nothing Then Err.Clear : Exit Function
    Dim j As Long
    For j = 1 To bs.Count
        Dim b As Object
        Set b = bs.Item(j)
        Dim t As String
        Dim n As String
        t = CStr(b("text"))
        n = UCase$(Trim$(Replace(t, "&", "")))
        If n = "DEBUG" Then WoMgHasDebugButton = True
        If InStr(1, n, UCase$(ChrW$(&H8C03) & ChrW$(&H8BD5))) > 0 Then WoMgHasDebugButton = True
        If WoMgHasDebugButton Then Exit Function
    Next
    Err.Clear
End Function

Private Sub WoMgRecordDialog(ByVal d As Object, ByVal category As String, ByVal action As String, ByVal clickedId As Long, ByVal clickedText As String, Optional ByVal plannedAction As String = "", Optional ByVal recordResult As String = "")
    On Error Resume Next
    If m_mgDialogs Is Nothing Then Set m_mgDialogs = New Collection
    Dim e As Object
    Set e = CreateObject("Scripting.Dictionary")
    e("hwnd") = CLng(d("hwnd"))
    e("cls") = CStr(d("cls"))
    e("title") = CStr(d("title"))
    e("bodyHead") = Left$(WoMgBodyOf(d), 200)
    e("ownerDisabled") = CBool(d("ownerDisabled"))
    Dim btns As Object
    Set btns = New Collection
    Dim bs As Object
    Set bs = btnsOf(d)
    If Not bs Is Nothing Then
        Dim j As Long, b As Object, be As Object
        For j = 1 To bs.Count
            Set b = bs.Item(j)
            Set be = CreateObject("Scripting.Dictionary")
            be("caption") = CStr(b("text"))
            be("id") = CLng(b("id"))
            be("hwnd") = CLng(b("hwnd"))
            btns.Add be
        Next
    End If
    Set e("buttons") = btns
    e("classification") = category
    e("action") = action
    ' FIX-21 (B1): every ledger entry carries the PLANNED action and an entry-level result, and
    ' held is DERIVED from them so it can never contradict the action.
    If Len(plannedAction) = 0 Then plannedAction = action
    e("plannedAction") = plannedAction
    If Len(recordResult) = 0 Then
        If LCase$(plannedAction) = "hold" Or InStr(1, action, "hold", vbTextCompare) > 0 Then
            recordResult = "held_policy"
        ElseIf clickedId <> 0 Then
            recordResult = "clicked"
        Else
            recordResult = "none"
        End If
    End If
    e("result") = recordResult
    ' FIX-22d: ONE variable decides held, so action/plannedAction/result/held cannot disagree.
    Dim heldFlag As Boolean
    heldFlag = False
    If LCase$(plannedAction) = "hold" Then heldFlag = True
    If InStr(1, action, "hold", vbTextCompare) > 0 Then heldFlag = True
    If LCase$(recordResult) = "held_policy" Then heldFlag = True
    e("held") = heldFlag
    e("heldBasis") = "plannedAction=" & plannedAction & "|result=" & recordResult
    ' 'my held box is still there' is distinguishable from 'no box'.
    If heldFlag Then
        e("heldBoxStillPresent") = (WoIsWindow(CLng(d("hwnd"))) <> 0)
    End If
    e("clickedId") = clickedId
    e("clickedCaption") = clickedText
    e("latencyMs") = IIf(m_mgAt = 0, 0, WoGetTickCount() - m_mgAt)
    e("source") = "modal_guard"
    e("preExisting") = WoMgIsPre(CLng(d("hwnd")))
    ' FIX-18: keep the structural evidence in the ledger, not only in the classifier.
    If d.Exists("sessionMsgBox") Then e("sessionMsgBox") = CBool(d("sessionMsgBox"))
    If d.Exists("programMsgBoxEvidence") Then e("programMsgBoxEvidence") = CStr(d("programMsgBoxEvidence"))
    If InStr(1, action, "hold", vbTextCompare) > 0 Or category = "program_msgbox" Then
        If Not e.Exists("sessionMsgBox") Then e("sessionMsgBox") = (category = "program_msgbox")
        ' FIX-22h: NEVER narrow held here. A program_msgbox record whose plannedAction=hold /
        ' result=held_policy already has held=True from heldFlag; the old assignment overwrote it
        ' with False because the ACTION name ("preexisting_held") does not contain "hold".
        e("held") = CBool(e("held")) Or (InStr(1, action, "hold", vbTextCompare) > 0)
    End If
    m_mgDialogs.Add e
    Err.Clear
End Sub

Public Function WoModalDialogs() As Object
    If m_mgDialogs Is Nothing Then Set m_mgDialogs = New Collection
    Set WoModalDialogs = m_mgDialogs
End Function

Public Function WoModalGuardAwaiting() As Boolean
    WoModalGuardAwaiting = m_mgAwaiting
End Function

Public Function WoModalGuardHowToRecover() As String
    If m_mgAwaiting Or m_mgUnresolvable Then
        WoModalGuardHowToRecover = "the in-call guard HELD this dialog (policy: " & m_mgCategory & "/" & m_mgAction & "). From ANOTHER connection: read_active_dialogs (read-only) -> send_window_message{hwnd, action:""click"", message:""bm_click""} on a SAFE caption (OK/Cancel/Close; NEVER End/Continue) -> ide_state{release_inflight:true, release_min_held_ms:<ms>} to free this call's single-flight gate."
    Else
        WoModalGuardHowToRecover = ""
    End If
End Function

' One guard pass. Idempotent, re-entrancy guarded, and self-terminating on its own
' deadline so a caller that never comes back cannot leave the guard armed for ever
' (which would keep WoConfirmTick standing down).
Public Function WoModalGuardTick() As Boolean
    On Error Resume Next
    WoModalGuardTick = False
    If Not m_mgArmed Then Exit Function
    If m_mgBusy Then Exit Function
    If m_woScanBusy Then Exit Function
    m_mgTickCount = m_mgTickCount + 1
    ' FIX-11: a REFUSED/HOLD decision no longer stops the tick - the hard deadline below
    ' still has to run, otherwise 'seen but held' wedges the in-flight call for ever.
    If WoGetTickCount() >= m_mgDeadline Then
        If m_mgSeen And Not m_mgClicked Then m_mgTimedOut = True
        m_mgArmed = False
        Err.Clear
        Exit Function
    End If
    If m_mgClicked Then
        ' the press already happened: this pass only checks whether it worked.
        m_mgTries = m_mgTries + 1
        If WoIsWindow(m_mgHwnd) = 0 Then
            m_mgDismissed = True
            m_mgArmed = False
        End If
        WoModalGuardTick = True
        Err.Clear
        Exit Function
    End If
    m_mgBusy = True
    m_mgTries = m_mgTries + 1
    Dim ds As Object
    Set ds = WoDialogsJson("#32770", 0, True)
    Dim n As Long
    n = 0
    If Not ds Is Nothing Then n = ds.Count
    If n = 0 Then
        m_mgBusy = False
        Err.Clear
        Exit Function
    End If
    m_mgSeen = True
    m_mgSeenCount = n
    If m_mgFirstSeenAt = 0 Then m_mgFirstSeenAt = WoGetTickCount()
    If m_mgButtons Is Nothing Then
        If n > 0 Then Set m_mgButtons = WoMgButtonsOf(ds.Item(1))
    End If
    ' FIX-11: HARD DEADLINE. The box has blocked this call past MODAL_WEDGE_DEADLINE_MS:
    ' stop waiting and do ONE deterministic thing (safe button -> one-shot WM_CLOSE ->
    ' release the single-flight gate from the tick). WM_CLOSE is parameterless and only
    ' ever sent to THIS #32770 (never WM_QUIT / WM_SYSCOMMAND).
    If WoMgDeadlineExceeded() Then
        If CBool(ds.Item(1)("ownerDisabled")) Then
            Call WoMgForceUnstick(ds.Item(1))
            m_mgBusy = False
            Err.Clear
            Exit Function
        End If
    End If
    ' FIX-7 rev: ABSOLUTE lifetime cap. The per-sighting deadline renewal below keeps
    ' the guard eligible while the box is on screen, which is right for a box it can
    ' press - but for a box it can NEVER press it would stay armed for ever and the
    ' in-flight call would never be released except by the gate backstop. After twice
    ' the budget the guard gives up and says so (modalUnresolvable + howToRecover).
    If (WoGetTickCount() - m_mgFirstSeenAt) > (m_mgWaitMs * 2) Then
        m_mgUnresolvable = True
        m_mgTimedOut = True
        m_mgMethod = "none (unresolvable box: no caption or standard ID this guard may press)"
        Call WoMgWedgeAbort("absolute_cap: " & m_mgCategory)
        m_mgArmed = False
        m_mgBusy = False
        Err.Clear
        Exit Function
    End If
    ' (#4 + #12): the deadline is -this many ms after the LAST SIGHTING-,
    ' NOT -this many ms after the arm-. Measured on arm 32F7D24C : on the
    ' make_project path the host Make Project box appears only AFTER the old 15 s window
    ' had already expired, so EVERY tick left through the deadline branch above and the box
    ' was never dismissed (the call then died on the gate backstop with no receipt).
    ' Renewing here means the guard stays eligible for as long as the box is actually on
    ' screen; the deadline is now only a safety net for a box that appears and never returns.
    m_mgDeadline = WoGetTickCount() + m_mgWaitMs
    Dim i As Long
    For i = 1 To n
        Dim d As Object
        Set d = ds.Item(i)
        ' ONLY a box that DISABLES ITS OWNER blocks the IDE. A merely visible window
        ' (a modeless toolbox, a floating pane) must never be clicked.
        If CBool(d("ownerDisabled")) Then
            Dim mgRisk As String, mgReason As String, mgClass As String
            WoHostModalRiskOfDialog d, mgRisk, mgReason, mgClass
            Dim bdy As String
            bdy = WoMgBodyOf(d)
            Dim bang As Boolean, bangHit As String
            bang = WoDialogBodyDangerous(d, bangHit)
            If m_mgHwnd = 0 Then
                m_mgTitle = CStr(d("title"))
                m_mgHwnd = CLng(d("hwnd"))
                m_mgCls = CStr(d("cls"))
                m_mgBody = bdy
                m_mgRisk = mgRisk
                m_mgRiskClass = mgClass
                m_mgRiskReason = mgReason
            End If
            Dim mgCat As String
            Dim mgAct As String
            WoMgClassify d, mgCat, mgAct, WoMgIsPre(CLng(d("hwnd")))
            m_mgCategory = mgCat
            m_mgAction = mgAct
            Dim pick As Long, pickText As String, pickWhy As String, pickCancel As Boolean
            Dim pickId As Long
            If mgAct = "hold" Then
                m_mgAwaiting = True
                m_mgRefused = m_mgRefused + 1
                m_mgRefusedReason = "policy_hold: " & mgCat
                m_mgMethod = "none (policy holds this box: awaiting_user)"
                pickWhy = "policy_hold: " & mgCat
                WoMgRecordDialog d, mgCat, mgAct, 0, "", mgAct, "held_policy"
                pick = 0
                ' FIX-11: a box this CALL opened is blocking the call right now. Releasing
                ' the single-flight gate from the tick is the only thing our code can do.
                If Not WoMgIsPre(CLng(d("hwnd"))) Then
                    Call WoMgWedgeAbort("policy_hold: " & mgCat)
                End If
            Else
            pick = WoMgPickForAction(d, mgAct, pickText, pickWhy, pickCancel, pickId)
            If pick <> 0 Then
                If mgRisk = "known" And Not pickCancel And Not m_mgAllowHostModal Then
                    ' VB6's own Make Project box: only a cancel-like caption may be
                    ' pressed here. Pressing "the first safe button" starts a compile
                    ' inside the host's modal loop = defect #4.
                    m_mgRefused = m_mgRefused + 1
                    m_mgRefusedReason = "host_modal_risk_known: only a cancel-like caption is pressed on this box (pressing its confirm caption starts a compile inside the host modal loop)"
                    m_mgMethod = "none (refused before any message was sent)"
                ElseIf bang And Not pickCancel Then
                    m_mgRefused = m_mgRefused + 1
                    m_mgRefusedReason = "dangerous_dialog_body: " & bangHit
                    m_mgMethod = "none (refused before any message was sent)"
                Else
                    m_mgClickedHwnd = pick
                    m_mgClickedText = pickText
                    Dim vMth As String
                    Dim vRes As String
                    vMth = "BM_CLICK"
                    vRes = "ok"
                    Call WoMgPressVerify(d, pick, pickId, pickText, vMth, vRes)
                    m_mgMethod = vMth
                    Call WoMgRecordTick("click_safe", CLng(d("hwnd")), pickId, pickText, vMth, vRes)
                    ' BM_CLICK is documented to do nothing on an INACTIVE dialog. It is sent
                    ' SYNCHRONOUSLY to a window of our own thread, so re-probing the box NOW
                    ' is authoritative - and if it is still there, use the dialog's own
                    ' WM_COMMAND/BN_CLICKED path once, exactly like close_active_dialogs.
                    If WoIsWindow(m_mgHwnd) <> 0 Then
                        m_mgMethod = "BM_CLICK+WM_COMMAND"
                        WoSendMessageW CLng(d("hwnd")), WM_COMMAND, pickId, pick
                    End If
                    WoMgRecordDialog d, mgCat, mgAct, pickId, pickText, mgAct, vRes
                    m_mgClickedId = pickId
                    m_mgClicked = True
                    If WoIsWindow(m_mgHwnd) = 0 Then m_mgDismissed = True
                    m_mgResolvedMs = WoGetTickCount() - m_mgAt
                    m_mgArmed = False
                    m_mgBusy = False
                    WoModalGuardTick = True
                    Err.Clear
                    Exit Function
                End If
            Else
                m_mgRefused = m_mgRefused + 1
                m_mgRefusedReason = pickWhy
                m_mgMethod = "none (no caption or standard ID this guard is allowed to press)"
                ' FIX-11: nothing safe to press on a box THIS call opened -> never wait for
                ' ever: release the gate from the tick and report a wedge abort.
                If Not WoMgIsPre(CLng(d("hwnd"))) Then
                    Call WoMgWedgeAbort("no_safe_button: " & mgCat)
                End If
            End If
            End If
        End If
    Next
    m_mgBusy = False
    Err.Clear
End Function

' The machine-readable guard receipt. Nothing is invented: modalDetected=true means a
' blocking box WAS seen; clickedButton="" means nothing was pressed.
Public Function WoModalGuardReceipt() As Object
    On Error Resume Next
    If m_mgAt = 0 Then
        Set WoModalGuardReceipt = Nothing
        Err.Clear
        Exit Function
    End If
    Dim root As Object
    Set root = CreateObject("Scripting.Dictionary")
    root("armedBy") = m_mgBy
    root("modalDetected") = m_mgSeen
    root("modalCount") = m_mgSeenCount
    root("modalHwnd") = m_mgHwnd
    root("modalCls") = m_mgCls
    root("modalTitle") = m_mgTitle
    root("modalBody") = m_mgBody
    root("modalHostRisk") = m_mgRisk
    root("modalHostRiskClass") = m_mgRiskClass
    root("modalHostRiskReason") = m_mgRiskReason
    root("modalClickedButton") = m_mgClickedText
    root("modalClickedHwnd") = m_mgClickedHwnd
    root("modalClicked") = m_mgClicked
    root("modalClickMethod") = m_mgMethod
    Dim mbs As Object
    Set mbs = New Collection
    If Not m_mgButtons Is Nothing Then
        Dim mbi As Long
        For mbi = 1 To m_mgButtons.Count
            mbs.Add m_mgButtons(mbi)
        Next
    End If
    Set root("modalButtons") = mbs
    root("modalCategory") = m_mgCategory
    root("modalAction") = m_mgAction
    root("modalAwaitingUser") = m_mgAwaiting
    root("modalCloseSent") = m_mgCloseSent
    root("tickCount") = m_mgTickCount
    root("firstSeenAtMs") = m_mgFirstSeenAt
    root("deadlineAtMs") = IIf(m_mgFirstSeenAt = 0, 0, m_mgFirstSeenAt + MODAL_WEDGE_DEADLINE_MS)
    root("refusedTicks") = m_mgRefused
    Dim lta As Object
    Set lta = New Collection
    If Not m_mgLastTickAction Is Nothing Then lta.Add m_mgLastTickAction
    If lta.Count > 0 Then
        Set root("lastTickAction") = m_mgLastTickAction
    Else
        root("lastTickAction") = "none"
    End If
    Dim mdl As Object
    Set mdl = New Collection
    If Not m_mgDialogs Is Nothing Then
        Dim mdi As Long
        For mdi = 1 To m_mgDialogs.Count
            mdl.Add m_mgDialogs(mdi)
        Next
    End If
    Set root("dialogs") = mdl
    Dim wj As Object
    Set wj = WoModalWedgeJson()
    If Not wj Is Nothing Then Set root("wedgeAborted") = wj
    Dim recov As Object
    Set recov = WoModalRecovery()
    If Not recov Is Nothing Then Set root("recovery") = recov
    root("modalUnresolvable") = m_mgUnresolvable
    If m_mgUnresolvable Then
        root("howToRecover") = "the box that appeared offers NO caption and NO standard control ID this guard is allowed to press (or it is an End/Continue-class runtime-error box, which the guard must never press). From ANOTHER connection: read_active_dialogs (read-only), then send_window_message{hwnd, action:""click"", message:""bm_click""} on a SAFE caption (OK/Cancel/Close - never End/Continue), then ide_state{release_inflight:true, release_min_held_ms:<ms>} to free the single-flight gate this call still holds."
    Else
        root("howToRecover") = ""
    End If
    root("modalDismissed") = m_mgDismissed
    root("modalResolvedMs") = m_mgResolvedMs
    root("modalRefusedCount") = m_mgRefused
    root("modalRefusedReason") = m_mgRefusedReason
    root("modalAllowHostModal") = m_mgAllowHostModal
    root("modalWaitMs") = m_mgWaitMs
    root("modalWaitedMs") = IIf(m_mgAt = 0, 0, WoGetTickCount() - m_mgAt)
    root("modalTicks") = m_mgTries
    root("modalPolicy") = "CANCEL-FIRST, bounded. An exact Cancel/Close/Ignore/No caption is always preferred; a non-cancel caption is pressed ONLY when the dialog is not host-modal risk=known and its body names no destructive action. VB6's Make Project box enumerates Open(read-only,disabled)/OK(id1)/Cancel(id2)/Help/Options, so 'the first safe button' would press OK and start a compile inside the host's modal loop - the guard never does that."
    Dim rb As String
    rb = "not_needed"
    If m_mgSeen Then
        If m_mgClicked Then
            rb = "modal_guard_click"
        Else
            rb = "modal_guard_timeout"
        End If
    End If
    If m_mgTimedOut Then rb = "modal_guard_timeout"
    root("modalRecoveredBy") = rb
    Dim ses As String
    If m_mgClicked Then
        If m_mgDismissed Then
            ses = "none: caption '" & m_mgClickedText & "' was pressed and the box is gone. The guard never presses a CONFIRM caption on a host-modal box, so it did not start that box's operation."
        Else
            ses = "unknown: caption '" & m_mgClickedText & "' was pressed but the box was still open at the last check - re-probe with read_active_dialogs"
        End If
    ElseIf m_mgSeen Then
        ses = "none: nothing was sent (see modalRefusedReason)"
    Else
        ses = "none: no blocking dialog was seen inside the budget"
    End If
    root("modalSideEffects") = ses
    root("modalNote") = "the guard is driven by the watchdog tick (ServerForm.tmrWatch -> WheelTick), the only channel the host's own modal loop still dispatches. It is a bounded SAFETY NET, not a proof: modalDetected=false means no blocking box was seen within modalWaitMs, NOT that none existed."
    Set WoModalGuardReceipt = root
    Err.Clear
End Function
Private Function btnsOf(ByVal d As Object) As Object
    On Error Resume Next
    Set btnsOf = d("buttons")
End Function

' =====================================================================
' (1): send_window_message
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
    ' (contract v1.6.0): HIGH-RISK message guard. WM_CLOSE / WM_QUIT /
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
' optional code-pane wheel - OPT-IN, DEFAULT OFF.
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
    ' record EVERY observed WM_MOUSEWHEEL with its raw wParam/lParam, so the
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
    ' final fix: move the pane through the VBIDE object model.
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
        ' (the code-pane caption measurement): "the pane did not resolve" and "the pane resolved but the
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

', LEGACY (no longer the behaviour oracle as of): first
' visible line of the pane's TEXT child. EM_GETFIRSTVISIBLELINE is the RichEdit
' message (0xCE); measured "RichEdit20A" as the code pane's text
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

' which of the three candidate reasons the CURRENT wheel state points at.
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

' READ-ONLY wheel snapshot, embedded as the "wheel" object of mcp_info and
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
    d("envExplicitOff") = m_whEnvOff
    d("periodicTimerArmed") = (m_whFastId <> 0)
    d("attachRequests") = m_whAttachReqs
    d("attachChecks") = m_whAttachChecks
    d("attachHits") = m_whAttachHits
    d("attachMisses") = m_whAttachMisses
    d("attachRetries") = m_whAttachRetries
    d("attachMaxRetries") = m_whAttachMaxRetries
    If m_whEnvOff Then
        d("wheelMode") = "env_off"
    ElseIf m_whEnabled Then
        d("wheelMode") = "on"
    Else
        d("wheelMode") = "off"
    End If
    d("attachSuppressed") = m_whAttachSuppressed
    d("attachGapMs") = WHEEL_ATTACH_GAP_MS
    d("focSeen") = m_whFocSeen
    d("lastScanAt") = m_whLastScanAt
    d("foc") = m_whFocHwnd
    d("focBasis") = m_whFocBasis
    d("hooked") = m_whFocHooked
    d("focusChecks") = m_whFocusChecks
    ' R2 WHEEL-FAST counters (additive; the old fields above are unchanged):
    '   fastSkips    = ticks the event gate suppressed (the CPU this patch removes)
    '   focusEvents  = SetWinEventHook callbacks received
    '   hookFallback = 1 when SetWinEventHook failed and the per-tick probe is back
    '   fastOn       = the runtime switch (VB6MCP_WHEEL_FAST, default 1)
    d("fastSkips") = m_whFastSkips
    d("focusEvents") = m_whFocusEvents
    d("hookFallback") = m_whHookFallback
    d("fastOn") = m_whFastOn
    ' R2 teardown proof, readable at ANY time while the DLL is loaded:
    '   evtHook  = the live WinEvent hook handle (0 = none installed)
    '   unhooked = (evtHook = 0) - the read-back a lane can quote at close
    '   unhookCalls / unhookOk = removal attempts / proven removals
    d("evtHook") = m_whEvtHook
    d("unhooked") = (m_whEvtHook = 0)
    d("unhookCalls") = m_whUnhookCalls
    d("unhookOk") = m_whUnhookOk
    d("attachReasonRpc") = m_whReasonRpc
    d("attachReasonFocus") = m_whReasonFocus
    d("attachReasonArm") = m_whReasonArm
    d("attachReasonEvent") = m_whReasonEvent
    d("attachReasonRetry") = m_whReasonRetry
    d("attachReasonOther") = m_whReasonOther
    d("scans") = m_whScans
    d("skippedAlreadyHooked") = m_whSkipAlready
    d("hooksEverInstalled") = m_whHooksInstalled
    d("shutdownCalls") = m_whShutdownCalls
    d("installed") = m_whInstalled
    d("armed") = (m_whCount > 0)
    d("subclassCount") = m_whCount
    ' every hooked window IS a code pane now (WheelScan refuses any
    ' VbaWindow whose title is not a live CodePane caption), so these fields can no
    ' longer point at the Immediate window - the misleading observability that sent
    ' an earlier measurement to the wrong window.
    d("targetHwnd") = 0
    d("targetTitle") = ""
    d("targetCaption") = ""
    ' the FIRST CODE PANE, not merely the first attached window (the fast
    ' attach now subclasses the Immediate/Watch/Locals windows too, and reads
    ' target* as "which CODE PANE is armed" - a non-pane window must never appear here).
    Dim ti As Long
    For ti = 0 To m_whCount - 1
        If m_whIsPane(ti) Then
            d("targetHwnd") = m_whHwnd(ti)
            d("targetTitle") = m_whTitle(ti)
            d("targetCaption") = m_whTitle(ti)
            Exit For
        End If
    Next
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
    ' (the pane-scroll measurement): -1 must never look like "line 0" or like "did not move".
    ' (the pane-scroll measurement): -1 must never look like "line 0" or like "did not move".
    d("lastTopLineBeforeKnown") = (m_whLastTopBefore > 0)
    d("lastTopLineBeforeNote") = ""
    If m_whLastTopBefore < 0 Then
        d("lastTopLineBeforeNote") = "UNAVAILABLE, not line 0. The exact cause is in lastTopStatus/lastTopError: pane_not_found = m_whLastTitle did not resolve to a CodePane; pane_not_writable_in_break = the pane WAS resolved but the IDE refused the TopLine read/write (measured 2026-10-13 in break mode; the real VB error is carried in lastTopError). Read paneTopLineNow (ActiveCodePane.TopLine, stamped by CodeTools) for the CURRENT line instead of comparing this field."
    End If
    ' (the code-pane caption measurement): why the caption table was empty/partial on the last scan.
    d("captionsSource") = m_whCapsSource
    d("captionsCount") = m_whCapsCount
    d("lastResolveVia") = m_whResolveVia
    d("lastTopStatus") = m_whTopStatus
    d("lastTopError") = m_whTopErr
    d("pendingScroll") = m_whPendOn
    d("pendingScrollResult") = m_whPendResult
    d("topLineSpace") = "lastTopLineBefore/After are VBIDE CodePane.TopLine (1-based) read back around the scroll; -1 means NO VALUE and the reason is in lastTopStatus/lastTopError: pane_not_found = the title resolved to no CodePane (forwarded untouched), pane_not_writable_in_break = the pane WAS resolved but the IDE refused the TopLine read/write (real VB error in lastTopError; the scroll is retried outside the window procedure by the timer - see pendingScroll/pendingScrollResult). The old EM_GETFIRSTVISIBLELINE channel is gone from these two fields because the code pane exposes no RichEdit20A child on this IDE (it read -1 and was never the behaviour oracle). paneTopLineNow is the ActiveCodePane.TopLine stamped by CodeTools."
    d("lastReason") = m_whLastReason
    d("reason") = WheelReason()
    ' hooked[] keeps its meaning (the subclassed CODE PANES only).
    ' attached[] is the new, complete list of every subclassed window with its role, and
    ' attachDecisions[] (option E) says why each VbaWindow seen in the last pass was or
    ' was not hooked - so a missing hook can never again look like a missing window.
    Dim hs As Object
    Set hs = New Collection
    Dim att As Object
    Set att = New Collection
    Dim paneN As Long
    paneN = 0
    For i = 0 To m_whCount - 1
        Dim he As Object
        Set he = CreateObject("Scripting.Dictionary")
        he("hwnd") = m_whHwnd(i)
        he("cls") = WoCls(m_whHwnd(i))
        he("title") = WoText(m_whHwnd(i))
        he("caption") = m_whTitle(i)
        he("isCodePane") = m_whIsPane(i)
        att.Add he
        If m_whIsPane(i) Then
            hs.Add he
            paneN = paneN + 1
        End If
    Next
    Set d("hooked") = hs
    Set d("attached") = att
    d("attachedCount") = m_whCount
    d("codePaneHookedCount") = paneN
    Dim ds As Object
    Set ds = New Collection
    Dim di As Long
    For di = 0 To m_whDecCount - 1
        Dim de As Object
        Set de = CreateObject("Scripting.Dictionary")
        de("hwnd") = m_whDecHwnd(di)
        de("title") = m_whDecTitle(di)
        de("decision") = m_whDecDecision(di)
        ds.Add de
    Next
    Set d("attachDecisions") = ds
    d("fastAttachIntervalMs") = WHEEL_FAST_MS
    d("fastAttachId") = m_whFastId
    d("fastAttachTicks") = m_whFastTicks
    d("fastAttachSkipped") = m_whFastSkipped
    d("fastAttachSafetyMs") = WHEEL_FAST_SAFETY_MS
    d("fastAttachGate") = ": the fast pass runs only when the FOREGROUND hwnd changed, or after WHEEL_FAST_SAFETY_MS since the last pass; every other 250 ms tick increments fastAttachSkipped instead of enumerating the window tree."
    d("attachRuns") = m_whAttachRuns
    d("lastAttachMsAgo") = IIf(m_whFastLastAt = 0, -1, WoGetTickCount() - m_whFastLastAt)
    d("attachNote") = "hooked[] lists the subclassed CODE PANES. attached[] lists EVERY subclassed window (the Immediate/Watch/Locals windows are subclassed too now, and are forwarded untouched at message time). attachDecisions[] says why each VbaWindow of the last pass was or was not hooked. fastAttachTicks/attachRuns show the COM-free attach timer that closes the gap between a pane being created and the 5 s watchdog scan."
    Set WheelStateObj = d
End Function

' Handle counts for the leak budget: GDI (flag 0) and USER (flag 1) objects of this
' process. Read-only; the host logs them at start and at shutdown so a slow leak shows.
' Leak counters for the OOM triage table. One call answers ALL FOUR columns:
'   gdi / user      GDI+USER objects of this process (flag 0/1)
'   handleCount     kernel handles (GetProcessHandleCount)
'   privateMb       commit charge (PROCESS_MEMORY_COUNTERS.PagefileUsage)
'   workingSetMb    resident set
'   virtualMb       sum of COMMITTED address-space regions (VirtualQuery walk): a
'                   32-bit VB6.EXE has no LAA, so exhaustion shows up here first
' Read-only. All fields are Null when their API failed (never a fabricated zero).
Public Function WheelGuiResourcesJson() As String
    On Error Resume Next
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim hp As Long
    Dim hc As Long
    hp = WoGetCurrentProcess()
    d("gdi") = WoGetGuiResources(hp, 0)
    d("user") = WoGetGuiResources(hp, 1)
    hc = 0
    If WoGetProcessHandleCount(hp, hc) <> 0 Then d("handleCount") = hc
    ' PROCESS_MEMORY_COUNTERS (9 dwords) is enough for WorkingSetSize + PagefileUsage.
    Dim mc(0 To 9) As Long
    mc(0) = LenB(mc(0)) * 10   ' cb
    If WoGetProcessMemoryInfo(hp, mc(0), LenB(mc(0)) * 10) <> 0 Then
        d("workingSetMb") = CLng(mc(3) \ 1048576)
        d("privateMb") = CLng(mc(8) \ 1048576)
    End If
    ' Address-space walk. COMMITTED sum is 'how much is in use'; the FREE-region numbers
    ' are the real error-7 predictor on a 32-bit host (VB6.EXE has no LAA): a process can
    ' be far from 2 GB committed and still fail a large allocation because the LARGEST
    ' contiguous free block collapsed.
    Dim mbi(0 To 6) As Long
    Dim addr As Long
    Dim q As Long
    Dim total As Double
    Dim freeTotal As Double
    Dim largestFree As Double
    Dim committedRegions As Long
    Dim freeRegions As Long
    addr = 0
    total = 0
    freeTotal = 0
    largestFree = 0
    committedRegions = 0
    freeRegions = 0
    Do While addr >= 0
        q = WoVirtualQuery(addr, mbi(0), LenB(mbi(0)) * 7)
        If q = 0 Then Exit Do
        If mbi(3) <= 0 Then Exit Do
        If mbi(4) = WoMEM_COMMIT Then
            total = total + CDbl(mbi(3))
            committedRegions = committedRegions + 1
        ElseIf mbi(4) = WoMEM_FREE Then
            freeTotal = freeTotal + CDbl(mbi(3))
            freeRegions = freeRegions + 1
            If CDbl(mbi(3)) > largestFree Then largestFree = CDbl(mbi(3))
        End If
        If CDbl(addr) + CDbl(mbi(3)) > 2147483647# Then Exit Do
        addr = addr + mbi(3)
    Loop
    d("virtualMb") = CLng(total / 1048576#)
    d("freeMb") = CLng(freeTotal / 1048576#)
    d("largestFreeMb") = CLng(largestFree / 1048576#)
    d("committedRegions") = committedRegions
    d("freeRegions") = freeRegions
    d("largestFreeNote") = "largestFreeMb is the largest CONTIGUOUS MEM_FREE block: below ~32 MB a large allocation can raise error 7 even with plenty of committed headroom; below ~8 MB it is effectively out of address space"
    d("queryOk") = True
    WheelGuiResourcesJson = JsonStringify(d)
    Err.Clear
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

' h may only be hooked when cap is a live CodePane caption (see
' WheelScan). The caption is stored next to the hwnd because it IS the key the
' scroll needs - and it is read here, in the scan, never inside the WndProc
' (enumeration/COM in a window procedure is a re-entrancy hazard).
Private Sub WheelHook(ByVal h As Long, ByVal cap As String, Optional ByVal isPane As Boolean = False)
    On Error Resume Next
    If h = 0 Then Exit Sub
    If Len(cap) = 0 Then Exit Sub
    If WheelHooked(h) Then
        m_whSkipAlready = m_whSkipAlready + 1
        Exit Sub
    End If
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
    ' remember WHETHER the table classified this window as a code pane,
    ' so hooked[] can keep the meaning while attached[] shows everything.
    m_whIsPane(m_whCount) = isPane
    m_whCount = m_whCount + 1
    m_whInstalled = True
    m_whHooksInstalled = m_whHooksInstalled + 1
End Sub

' is cap the caption of a code pane that is open RIGHT NOW?
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

' exactly one VbaWindow is a code pane; everything else with the
' SAME class (the Immediate / Watch / Locals windows) is skipped and recorded,
' so the receipt can never again show a non-code window as the target.
'
' (measured 2026-10-13, pre-fix build D454878E): the VBE caption table can come
' back EMPTY while the project is running/broken - the debugger opens the code pane of the
' module it stopped in, and that pane's window need never have existed during design. In
' break mode the watchdog KEPT SCANNING (scans 2 -> 5) yet WheelPaneCaptions() yielded 0
' captions, so every VbaWindow - including the pane showing the current statement - was
' recorded "not_a_code_pane" and never subclassed; a WM_MOUSEWHEEL posted to that pane never
' reached our WndProc (msgSeen did not move) and the wheel did nothing. With nothing to
' classify against, refusing to hook IS the defect. The scroll path re-resolves the caption
' at message time and FORWARDS anything that does not resolve to a CodePane, so hooking an
' unclassified VbaWindow cannot change that window's own behaviour. With a NON-empty table
' the old strict rule is unchanged (criterion 4 evidence still holds).
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
        m_whEverVerified = True
        WheelMarkPane h
        WheelHook h, cap, True
        WheelDecide h, cap, "classified_code_pane"
    ElseIf Len(cap) > 0 And m_whEverVerified And WheelTableEmpty(caps) Then
        WheelMarkPane h
        WheelHook h, cap, True
        WheelDecide h, cap, "empty_table_fallback"
    ElseIf Len(cap) > 0 Then
        ' (A): the caption table NO LONGER GATES THE HOOK. Measured on the
        ' live IDE: a code pane opened after the last scan is absent from the table until
        ' the next scan, and this scan is the only thing that hooks a window - so gating
        ' on the table meant a brand-new pane was never subclassed while the user scrolled
        ' (msgSeen did NOT move, so (i) "not attached", not (ii) "resolve failed").
        ' The window is subclassed regardless; the classification stays where it belongs
        ' - at message time - and WheelHandle still forwards anything that does not
        ' resolve to a CodePane untouched, so is preserved.
        WheelHook h, cap, False
        WheelDecide h, cap, "hooked_unclassified"
    Else
        WheelSkip h, cap, "not_a_code_pane"
        WheelDecide h, cap, "skip_empty_text"
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

' =====================================================================
' (wheel #1): the FAST, COM-FREE attach pass.
'   Why it exists: the only driver of WheelScan is the watchdog
'   (ServerForm.tmrWatch, Interval = 5000 in steady state - measured 2 scans / 10 s) plus
'   the HTTP message pump. A code pane opened after the last scan was therefore NOT
'   subclassed when the user's wheel arrived, and WM_MOUSEWHEEL never reached us at all.
'   Measured on the live IDE before this fix: a new pane's msgSeen did NOT move (i), and
'   the caption table DID list it again on the next scan (1 -> 2 -> 3) - so the defect was
'   the moment the hook was established, not a permanently incomplete table.
'   This pass enumerates and subclasses only: no VBIDE call, no COM, no allocation that
'   can raise. Classification stays at message time (WheelHandle) and in WheelScan.
' =====================================================================
Public Sub WheelAttachAll()
    On Error Resume Next
    If m_woScanBusy Then Exit Sub
    WheelAttachPass
    Err.Clear
End Sub

Private Sub WheelAttachPass()
    On Error Resume Next
    If Not m_whEnabled Then Exit Sub
    m_whAttachRuns = m_whAttachRuns + 1
    m_whFastLastAt = WoGetTickCount()
    Dim tops() As Long, nTops As Long, t As Long, k As Long
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
            If LCase$(WoCls(h)) = "vbawindow" Then WheelAttachOne h
            WoScanKids h
            For k = 0 To m_woCount - 1
                If LCase$(WoCls(m_woHwnd(k))) = "vbawindow" Then WheelAttachOne m_woHwnd(k)
            Next
        End If
    Next
    Err.Clear
End Sub

' Hook ONE window if it can carry the wheel. A window is refused only when we cannot even
' NAME it (empty caption) or when the hook table is full - NEVER because the caption
' table did not list it (that was exactly the defect).
Private Sub WheelAttachOne(ByVal h As Long)
    On Error Resume Next
    Dim cap As String
    cap = WoText(h)
    If Len(cap) = 0 Then
        WheelDecide h, cap, "skip_empty_text"
        Exit Sub
    End If
    If WheelHooked(h) Then Exit Sub
    If m_whCount >= WHEEL_MAX Then
        WheelDecide h, cap, "skip_hook_table_full"
        Exit Sub
    End If
    WheelHook h, cap, False
    WheelDecide h, cap, "hooked_by_fast_attach"
    Err.Clear
End Sub

Private Sub WheelDecide(ByVal h As Long, ByVal cap As String, ByVal decision As String)
    If m_whDecCount >= WHEEL_SKIP_MAX Then Exit Sub
    m_whDecHwnd(m_whDecCount) = h
    m_whDecTitle(m_whDecCount) = cap
    m_whDecDecision(m_whDecCount) = decision
    m_whDecCount = m_whDecCount + 1
End Sub

' Upgrade an ALREADY hooked window to "this is a code pane". Needed because the fast
' attach can hook a pane before the scan that builds the caption table has run.
Private Sub WheelMarkPane(ByVal h As Long)
    Dim i As Long
    For i = 0 To m_whCount - 1
        If m_whHwnd(i) = h Then m_whIsPane(i) = True
    Next
End Sub

' (B): the SetTimer callback. The OS calls it from the IDE's message loop,
' so it must never raise and must never touch the VBIDE object model - which is why the
' fast attach is COM-free by construction.
Public Sub WheelTimerProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    m_whFastTicks = m_whFastTicks + 1
    If Not m_whEnabled Then
        Err.Clear
        Exit Sub
    End If
    ' gate the expensive pass (see WHEEL_FAST_SAFETY_MS).
    Dim fastFg As Long, fastNow As Long
    fastFg = WoGetForegroundWindow()
    fastNow = WoGetTickCount()
    If fastFg = m_whFastFg Then
        ' wraparound-safe subtraction: never compute now + 5000 (Long overflow)
        If (fastNow - m_whFastLastAt) < WHEEL_FAST_SAFETY_MS Then
            m_whFastSkipped = m_whFastSkipped + 1
            Err.Clear
            Exit Sub
        End If
    End If
    m_whFastFg = fastFg
    WheelAttachAll
    Err.Clear
End Sub

' The 250 ms fast-attach timer is RETIRED. Pane discovery is event driven now: VBIDE
' project/component events call WheelRequestAttach, which runs ONE scan per event.
' Nothing periodic remains in the wheel, so idle costs exactly zero.
Private Sub WheelStartFast()
    m_whFastId = 0
    m_whFastIn = False
End Sub

Private Sub WheelStopFast()
    On Error Resume Next
    If m_whFastId <> 0 Then WoKillTimer 0, m_whFastId
    m_whFastId = 0
    Err.Clear
End Sub

' R2 WHEEL-FAST: arm the focus event channel. A zero handle is FAIL-OPEN: it is reported
' (wheel.hookFallback=1) and the tick keeps polling every tick, i.e. today's behaviour.
Private Sub WheelArmFocusHook()
    On Error Resume Next
    ' a hook must never be armed while we are dismantling (or after the host told us
    ' to go away): it would reintroduce the unmapped-callback crash shape we are removing.
    If m_whShutdown Then Exit Sub
    m_whEvtHook = 0
    m_whFocusDirty = True
    m_whFocusCheckedAt = 0
    m_whHookFallback = 0
    m_whFocusEvents = 0
    m_whFastSkips = 0
    If Not m_whFastOn Then Exit Sub
    m_whEvtHook = WoSetWinEventHook(WHEEL_EVT_OBJECT_FOCUS, WHEEL_EVT_OBJECT_FOCUS, 0&, AddressOf WoWheelFocusProc, 0&, 0&, 0&)
    If m_whEvtHook = 0 Then
        m_whHookFallback = 1
        Util.LogLine "wheel fasthook unavailable (SetWinEventHook=0): per-tick focus polling"
    End If
    Err.Clear
End Sub

' R2 WHEEL-FAST: retire the channel. Called from the unload path so no callback can
' outlive the DLL (the same rule the window subclass teardown follows).
Private Sub WheelDisarmFocusHook()
    ' READ-BACK, not a promise. UnhookWinEvent returns nonzero when the hook WAS removed.
    ' A SECOND call on the same handle must then return 0, because the handle is no longer
    ' registered - that zero is the read-back. If the first call failed, the second is a
    ' real retry: a nonzero second return means THIS call removed it. Both returns are
    ' logged and only 'both zero' counts as a failure (unhookOk stays put).
    On Error Resume Next
    Dim was As Long, rc1 As Long, rc2 As Long
    was = m_whEvtHook
    If was <> 0 Then
        m_whUnhookCalls = m_whUnhookCalls + 1
        rc1 = WoUnhookWinEvent(was)
        rc2 = WoUnhookWinEvent(was)
        If (rc1 <> 0 And rc2 = 0) Or (rc2 <> 0) Then m_whUnhookOk = m_whUnhookOk + 1
        m_whEvtHook = 0
    End If
    m_whFocusDirty = True
    Util.LogLine "wheel unhook was=" & CStr(was) & " rc1=" & CStr(rc1) & " rc2=" & CStr(rc2) & " now=" & CStr(m_whEvtHook) & " ok=" & CStr(m_whUnhookOk) & " calls=" & CStr(m_whUnhookCalls)
    Err.Clear
End Sub

' R2 WHEEL-FAST: the SetWinEventHook callback. It is dispatched on OUR thread inside the
' IDE pump, possibly between two messages of an in-flight tool call, so it must not
' create objects, must not log and must not touch COM/Swing: it raises a flag only.
Public Sub WoWheelFocusProc(ByVal hHook As Long, ByVal evt As Long, ByVal hwnd As Long, ByVal idObject As Long, ByVal idChild As Long, ByVal idThread As Long, ByVal msTime As Long)
    On Error Resume Next
    m_whFocusEvents = m_whFocusEvents + 1
    If m_whEnabled Then m_whFocusDirty = True
    Err.Clear
End Sub

' Drop entries whose window is gone, then hook every code pane that is not yet
' hooked. Hook the same hwnd twice is impossible (that is the no-stacking rule).
' EVENT-DRIVEN pane discovery. Called when the IDE's window/component structure may have
' changed (VBIDE item added/removed/renamed/reloaded, add-in start, a tool call). ONE scan
' per event; WheelHooked() refuses a window twice, so this is idempotent and idle-safe.
Public Sub WheelRequestAttach(ByVal reason As String)
    On Error Resume Next
    If m_whShutdown Then Exit Sub
    If Not m_whEnabled Then Exit Sub
    m_whAttachReqs = m_whAttachReqs + 1
    Select Case reason
    Case "rpc_begin", "rpc_end"
        m_whReasonRpc = m_whReasonRpc + 1
    Case "focus_unhooked"
        m_whReasonFocus = m_whReasonFocus + 1
    Case "arm"
        m_whReasonArm = m_whReasonArm + 1
    Case "ide_event"
        m_whReasonEvent = m_whReasonEvent + 1
    Case "deferred_retry"
        m_whReasonRetry = m_whReasonRetry + 1
    Case Else
        m_whReasonOther = m_whReasonOther + 1
    End Select
    ' RATE GATE: 'one scan per EVENT' must never degrade into 'one scan per tick' when
    ' an event keeps re-firing. 250 ms is far faster than a human can trigger; 'arm' is
    ' exempt because it is the single start-up scan.
    Dim gateNow As Long
    gateNow = WoGetTickCount()
    If reason <> "arm" Then
        If m_whLastScanAt <> 0 Then
            If (gateNow - m_whLastScanAt) < WHEEL_ATTACH_GAP_MS Then
                m_whAttachSuppressed = m_whAttachSuppressed + 1
                Err.Clear
                Exit Sub
            End If
        End If
    End If
    m_whLastScanAt = gateNow
    Dim before As Long
    before = m_whHooksInstalled
    WheelScan
    If m_whHooksInstalled > before Then
        m_whAttachHits = m_whAttachHits + 1
        m_whAttachPending = False
        m_whAttachNextAt = 0
        m_whAttachRetries = 0
    Else
        m_whAttachMisses = m_whAttachMisses + 1
        ' R2b: bounded ladder; it never re-arms itself (R2 renewed the deadline on
        ' every miss, so one ordinary 'nothing new to hook' pinned pending forever
        ' and the WHEEL_FAST exit could never be taken).
        If Not m_whAttachPending Then
            m_whAttachRetries = 0
            m_whAttachNextAt = WoGetTickCount() + WHEEL_ATTACH_RETRY_BASE_MS
            m_whAttachDeadline = WoGetTickCount() + WHEEL_ATTACH_LADDER_MS
            m_whAttachPending = True
        End If
    End If
    Err.Clear
End Sub

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
    ' (re)build the code-pane caption table for THIS scan. It is the
    ' ONLY thing that distinguishes a code pane from the Immediate / Watch / Locals
    ' windows, which carry the same window class "VbaWindow". Cost is one late-bound
    ' call per scan, not per message: nothing COM happens in the window procedure.
    Set m_whCaps = Nothing
    m_whSkipCount = 0
    If Not m_whSink Is Nothing Then
        Set m_whCaps = m_whSink.WheelPaneCaptions()
        ' (the code-pane caption measurement): publish WHY the table looks the way it does. In break
        ' mode the VBE.CodePanes collection yields nothing and CodeTools falls back to the
        ' active pane; that is now visible instead of looking like "no panes exist".
        m_whCapsSource = m_whSink.WheelPaneCaptionsSource()
    End If
    m_whCapsCount = 0
    If Not m_whCaps Is Nothing Then m_whCapsCount = m_whCaps.Count
    ' (A): subclass EVERY VbaWindow first, table or no table. The decision is
    ' kept out of the hook: an unclassified window is forwarded untouched at message time.
    m_woScanBusy = True
    ' the decision list belongs to the SCAN (the pass that classifies), not to
    ' the fast pass - a fast pass over an already-hooked set decides nothing and would wipe
    ' the list, which made the option-E diagnostic come back empty on the live IDE.
    m_whDecCount = 0
    WheelAttachPass
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
    m_woScanBusy = False
    Err.Clear
End Sub

' Throttled rescan; safe to call from the watchdog timer and the message pump.
' NOT a periodic scanner any more. It drains work that a MESSAGE or a VBIDE event already
' queued (modal-confirm click, modal guard, deferred pane scroll). With no events there is
' nothing to do, and no timer-driven work exists in the wheel any more.
Public Sub WheelTick()
    ' NOT a periodic scanner: these two drains only run work that a MESSAGE or an
    ' event already queued.
    Dim whT0 As Double, whPf As Boolean, whNow As Long
    whPf = WoProfOn()
    If whPf Then whT0 = WoProfNow()
    WoConfirmTick
    WoModalGuardTick
    WheelPumpPending
    If whPf Then WoProfAdd 17, WoProfNow() - whT0
    If m_whShutdown Then Exit Sub
    If Not m_whEnabled Then Exit Sub
    ' R2b: the deferred attach retry keeps its OWN bounded cadence and is serviced
    ' BEFORE the fast exit, so the fast path never depends on it.
    If m_whAttachPending Then
        Dim atNow As Long
        atNow = WoGetTickCount()
        If atNow >= m_whAttachDeadline Then
            m_whAttachPending = False
            m_whAttachNextAt = 0
            m_whAttachRetries = 0
        ElseIf m_whAttachRetries >= m_whAttachMaxRetries Then
            m_whAttachPending = False
            m_whAttachNextAt = 0
            m_whAttachRetries = 0
        ElseIf atNow >= m_whAttachNextAt Then
            m_whAttachRetries = m_whAttachRetries + 1
            m_whAttachNextAt = atNow + WHEEL_ATTACH_RETRY_BASE_MS * m_whAttachRetries
            WheelRequestAttach "deferred_retry"
        End If
    End If
    ' R2 WHEEL-FAST: an idle tick does NO Win32 call and never touches the hook table.
    ' The probe block below runs only when the SetWinEventHook callback marked it dirty.
    ' The 1000 ms branch is the correctness backstop, not the mechanism. The deferred
    ' attach retry keeps its own cadence, so it is never skipped by the fast exit.
    If m_whFastOn And m_whHookFallback = 0 Then
        whNow = WoGetTickCount()
        If Not m_whFocusDirty Then
            ' wraparound-safe subtraction: never compute now + 1000 (Long overflow)
            If (whNow - m_whFocusCheckedAt) >= 0 And (whNow - m_whFocusCheckedAt) < WHEEL_FOCUS_BACKSTOP_MS Then
                m_whFastSkips = m_whFastSkips + 1
                Exit Sub
            End If
        End If
        m_whFocusDirty = False
        m_whFocusCheckedAt = whNow
    End If
    ' O(1) attach trigger: the FOCUSED window is the pane the user would scroll. Looking
    ' it up costs two API calls (no window-tree walk); a hook is only installed when the
    ' focused window is NOT already in the hook table.
    ' CHANGE-DRIVEN, not state-driven. 'The focused window is not hooked' is the NORMAL
    ' state (focus sits on the Project Explorer or the main frame), so it must NEVER be
    ' read as 'scan'. One scan per focus CHANGE to a VbaWindow that is not hooked yet;
    ' an idle IDE does zero scans, and the gate in WheelRequestAttach bounds the rate.
    If whPf Then whT0 = WoProfNow()
    m_whAttachChecks = m_whAttachChecks + 1
    m_whFocusChecks = m_whFocusChecks + 1
    Dim whTb As Double
    If whPf Then whTb = WoProfNow()
    m_whFocHwnd = WoGetThreadFocus()
    If whPf Then WoProfAdd 23, WoProfNow() - whTb
    If m_whFocHwnd <> 0 Then
        If m_whFocHwnd <> m_whFocSeen Then
            m_whFocSeen = m_whFocHwnd
            m_whFocHooked = WheelHooked(m_whFocHwnd)
            If m_whFocHooked Then
                m_whAttachSuppressed = m_whAttachSuppressed + 1
            ElseIf LCase$(WoCls(m_whFocHwnd)) <> "vbawindow" Then
                m_whAttachSuppressed = m_whAttachSuppressed + 1
            Else
                If whPf Then whTb = WoProfNow()
                WheelRequestAttach "focus_changed"
                If whPf Then WoProfAdd 25, WoProfNow() - whTb
            End If
        Else
            m_whFocHooked = WheelHooked(m_whFocHwnd)
        End If
    End If
    If whPf Then WoProfAdd 18, WoProfNow() - whT0
End Sub

' (the code-pane caption measurement): perform a pending pane scroll OUTSIDE the subclassed WndProc.
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

' wire the VBIDE side in. WinOps.bas is a STANDARD module with no
' VBIDE reference; CodeTools is the only holder of the VBE, so McpServer hands it
' over here and the two scroll helpers are called late-bound. Called once at
' start-up (ServerHost.Start -> McpServer.Init), before WheelInit.
Public Sub WheelSetSink(ByVal o As Object)
    On Error Resume Next
    Set m_whSink = o
    Err.Clear
End Sub

' OPT-IN. DEFAULT OFF: with VB6MCP_WHEEL unset, nothing is subclassed at all.
' OPT-IN. DEFAULT OFF. TRI-STATE env: an EXPLICIT 0/off/no/false in VB6MCP_WHEEL turns the
' wheel OFF and does NOT fall back to the registry; only an UNSET variable falls back to
' HKCU\Software\VB6MCP\Wheel. The old code could only turn it ON (env=0 fell through to the
' registry), which made an on/off A/B measurement impossible.
Public Sub WheelInit()
    On Error Resume Next
    Dim s As String
    Dim envSet As Boolean
    envSet = (Len(Environ$("VB6MCP_WHEEL")) > 0)
    s = LCase$(Trim$(Environ$("VB6MCP_WHEEL")))
    m_whShutdown = False
    m_whEnvOff = (envSet And Not (s = "1" Or s = "true" Or s = "on" Or s = "yes"))
    If envSet Then
        m_whEnabled = (s = "1" Or s = "true" Or s = "on" Or s = "yes")
    Else
        s = LCase$(Trim$(WinOps.RegGetString("Software\VB6MCP", "Wheel")))
        m_whEnabled = (s = "1" Or s = "true" Or s = "on" Or s = "yes")
    End If
    m_whLines = 3
    s = Trim$(Environ$("VB6MCP_WHEEL_LINES"))
    If Len(s) = 0 Then
        s = Trim$(McpSettingGet("WheelLines"))
    End If
    If Len(s) > 0 Then
        If IsNumeric(s) Then
            If CLng(Val(s)) >= 1 And CLng(Val(s)) <= 50 Then m_whLines = CLng(Val(s))
        End If
    End If
    m_whFastOn = True
    s = LCase$(Trim$(Environ$("VB6MCP_WHEEL_FAST")))
    If s = "0" Or s = "off" Or s = "no" Or s = "false" Then m_whFastOn = False
    ' R2b: the retry ladder is bounded; 0 disables it entirely (pure event-driven
    ' attach) so the same binary can A/B the ladder without a rebuild.
    m_whAttachMaxRetries = WHEEL_ATTACH_MAX_RETRIES
    s = Trim$(Environ$("VB6MCP_WHEEL_ATTACH_RETRIES"))
    If Len(s) > 0 Then
        If IsNumeric(s) Then
            m_whAttachMaxRetries = CLng(Val(s))
            If m_whAttachMaxRetries < 0 Then m_whAttachMaxRetries = 0
            If m_whAttachMaxRetries > 4 Then m_whAttachMaxRetries = 4
        End If
    End If
    If m_whEnabled Then
'        ONE scan at arm time, counted as the 'arm' reason. After this the wheel is
'        message/event driven only (no periodic scan exists).
        WheelRequestAttach "arm"
        WheelArmFocusHook
    End If
    Err.Clear
End Sub

' CLEAN UNINSTALL. Restore the previous procedure wherever it is still ours;
' where another add-in subclassed after us, leave the chain intact and merely go
' inactive, because replacing the CURRENT procedure there would break THEIR
' chain. Reports what actually happened instead of claiming success.
' CLEAN UNINSTALL. Restore EVERY window procedure this add-in ever installed, then PROVE
' it by reading the procedure back. The old version restored only windows whose CURRENT
' procedure was still ours and merely went inactive when somebody had subclassed on top
' of us - but then OUR pointer is still reachable through THEIR chain, so the next message
' to that window ran unmapped code after the DLL was unloaded (measured: IDE-exit
' APPCRASH, faulting module VB6MCP.dll_unloaded, c0000005). Unloading therefore FORCES the
' original procedure back and reports exactly what it proved.
Public Function WheelShutdownJson() As String
    On Error Resume Next
    m_whEnabled = False
'    nothing may arm a new subclass while we are dismantling: the request latch is closed
    m_whShutdown = True
    m_whShutdownCalls = m_whShutdownCalls + 1
'    retire the (now unused) fast timer first - it must never fire during teardown
    WheelStopFast
    WheelDisarmFocusHook
    Dim restored As Long
    Dim verified As Long
    Dim forced As Long
    Dim skipped As Long
    Dim failed As Long
    Dim i As Long
    Dim cur As Long
    Dim okAll As Boolean
    Dim d As Object
    Dim fails As Object
    Dim e As Object
    restored = 0
    verified = 0
    forced = 0
    skipped = 0
    failed = 0
    okAll = True
    Set d = CreateObject("Scripting.Dictionary")
    Set fails = New Collection
    For i = 0 To m_whCount - 1
        If m_whHwnd(i) <> 0 Then
            If WoIsWindow(m_whHwnd(i)) = 0 Then
'                the window is already gone; no procedure is left in it
                skipped = skipped + 1
            ElseIf m_whPrev(i) = 0 Then
                failed = failed + 1
                okAll = False
                Set e = CreateObject("Scripting.Dictionary")
                e("hwnd") = m_whHwnd(i)
                e("why") = "no_saved_prev_proc"
                fails.Add e
            Else
                cur = WoGetWindowLongW(m_whHwnd(i), GWLP_WNDPROC)
                If cur = m_whSelf And m_whSelf <> 0 Then
                    WoSetWindowLongW m_whHwnd(i), GWLP_WNDPROC, m_whPrev(i)
                    restored = restored + 1
                Else
'                    somebody subclassed on top of us: OUR pointer is still inside THEIR
'                    chain, so it MUST come out or the unload crashes. Force ours back.
                    WoSetWindowLongW m_whHwnd(i), GWLP_WNDPROC, m_whPrev(i)
                    forced = forced + 1
                End If
'                READ-BACK: the only acceptable proof is the procedure we saved
                cur = WoGetWindowLongW(m_whHwnd(i), GWLP_WNDPROC)
                If cur = m_whPrev(i) Then
                    verified = verified + 1
                Else
                    failed = failed + 1
                    okAll = False
                    Set e = CreateObject("Scripting.Dictionary")
                    e("hwnd") = m_whHwnd(i)
                    e("why") = "readback_mismatch"
                    e("expected") = m_whPrev(i)
                    e("actual") = cur
                    fails.Add e
                End If
            End If
            m_whHwnd(i) = 0
            m_whPrev(i) = 0
            m_whTitle(i) = ""
            m_whIsPane(i) = False
        End If
    Next
    m_whCount = 0
    m_whInstalled = False
    m_whSelf = 0
    m_whEverVerified = False
    Set m_whSink = Nothing
    Set m_whCaps = Nothing
    m_whSkipCount = 0
    m_whPendOn = False
    d("wheelEnabled") = False
    d("calls") = m_whShutdownCalls
    d("ok") = okAll
    d("restored") = restored
    d("verified") = verified
    d("forced") = forced
    d("skippedWindowGone") = skipped
    d("failed") = failed
    d("hooksEverInstalled") = m_whHooksInstalled
    d("skippedAlreadyHooked") = m_whSkipAlready
    d("survivors") = m_whHooksInstalled - restored - forced - skipped
    d("balanced") = ((m_whHooksInstalled - restored - forced - skipped) = 0)
    Set d("failedDetail") = fails
    d("authority") = "every restored window is READ BACK with GetWindowLongW and must equal the saved original"
    d("note") = "no hook survives the unload; a failure here is reported, never silenced"
    ' the hook removal is proven by WheelDisarmFocusHook's read-back, so the receipt
    ' carries the proof itself (the full line is logged by ServerHost.Shutdown).
    d("evtHook") = m_whEvtHook
    d("unhooked") = (m_whEvtHook = 0)
    d("unhookCalls") = m_whUnhookCalls
    d("unhookOk") = m_whUnhookOk
    WheelShutdownJson = JsonStringify(d)
    Err.Clear
End Function

' (R18 focus safety): foreground read/write wrappers. They sit at the END of the
' module on purpose - VB6 only allows declarations before the first procedure, and this module
' has declaration blocks further down.
' FIX-17: deliver the VM's Toggle Breakpoint ACCELERATOR asynchronously. PostMessage returns
' immediately (the blocker was the synchronous ctrl.Execute), so this call can never wedge the
' tool. Reachability is proven BEFORE posting (same process + editor-class window); a false
' return means "channel unavailable", never a fake success.
' FIX-22: BOUNDED, message-free, COM-free liveness probe. Nothing here touches the VBIDE
' object model, so it is safe to call from the native observation loop while the IDE thread
' may be busy. respond=false is reported as-is (never guessed).
' FIX-22f: pure Win32 thread-state snapshot for the fire-and-forget native path. It uses
' IsHungAppWindow + GetClassName ONLY: no window message (no WM_NULL), no COM, no IDE call.
' The value is NOT the IDE design/break state and the receipt says so.
Public Function WoWin32ThreadState() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("state") = "unknown"
    d("mainHwnd") = 0
    d("mainClass") = ""
    d("hung") = False
    d("method") = "IsHungAppWindow+GetClassName (no message, no COM)"
    On Error Resume Next
    WoScanTops
    Dim n As Long
    n = m_woCount
    If n > WO_MAX Then n = WO_MAX
    Dim pid As Long
    pid = WoGetCurrentProcessId()
    Dim i As Long
    For i = 0 To n - 1
        Dim h As Long
        h = m_woHwnd(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                If WoIsWindowVisible(h) <> 0 Then
                    If LCase$(WoCls(h)) = "wndclass_desked_gsk" Then
                        d("mainHwnd") = h
                        d("mainClass") = WoCls(h)
                        d("hung") = (WoIsHungAppWindow(h) <> 0)
                        If CBool(d("hung")) Then d("state") = "hung" Else d("state") = "responsive"
                        Err.Clear
                        Set WoWin32ThreadState = d
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Err.Clear
    Set WoWin32ThreadState = d
End Function

Public Function WoHostLiveness(Optional ByVal timeoutMs As Long = 200) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("mainHwnd") = 0
    d("mainClass") = ""
    d("responds") = False
    d("hung") = False
    d("method") = "IsHungAppWindow+SendMessageTimeout(WM_NULL,SMTO_ABORTIFHUNG|SMTO_BLOCK)"
    d("timeoutMs") = timeoutMs
    Dim t0 As Long
    t0 = WoGetTickCount()
    On Error Resume Next
    WoScanTops
    Dim n As Long
    n = m_woCount
    If n > WO_MAX Then n = WO_MAX
    Dim pid As Long
    pid = WoGetCurrentProcessId()
    Dim hMain As Long
    hMain = 0
    Dim i As Long
    For i = 0 To n - 1
        Dim h As Long
        h = m_woHwnd(i)
        If h <> 0 Then
            If WoPidOf(h) = pid Then
                If WoIsWindowVisible(h) <> 0 Then
                    If LCase$(WoCls(h)) = "wndclass_desked_gsk" Then
                        hMain = h
                        Exit For
                    End If
                End If
            End If
        End If
    Next
    d("mainHwnd") = hMain
    If hMain <> 0 Then
        d("mainClass") = WoCls(hMain)
        d("hung") = (WoIsHungAppWindow(hMain) <> 0)
        Dim res As Long
        res = 0
        Dim sent As Long
        sent = WoSendMessageTimeoutW(hMain, WM_NULL, 0, 0, SMTO_ABORTIFHUNG Or SMTO_BLOCK, timeoutMs, res)
        d("responds") = (sent <> 0)
        d("ok") = True
    Else
        d("reason") = "ide_main_window_not_found"
    End If
    d("elapsedMs") = WoGetTickCount() - t0
    Err.Clear
    Set WoHostLiveness = d
End Function

Public Function WoPostF9ToFocus() As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("hwnd") = 0
    d("className") = ""
    d("reason") = "no_focus_window"
    On Error Resume Next
    Dim h As Long
    h = WoGetFocus()
    If h = 0 Then
        Err.Clear
        Set WoPostF9ToFocus = d
        Exit Function
    End If
    d("hwnd") = h
    Dim pid As Long
    pid = 0
    WoGetWindowThreadProcessId h, pid
    If pid <> WoGetCurrentProcessId() Then
        d("reason") = "foreign_process_focus"
        Err.Clear
        Set WoPostF9ToFocus = d
        Exit Function
    End If
    Dim buf As String
    Dim n As Long
    buf = String$(260, vbNullChar)
    n = WoGetClassNameW(h, StrPtr(buf), 260)
    If n > 0 Then d("className") = Left$(buf, n)
    Dim cn As String
    cn = CStr(d("className"))
    If InStr(1, cn, "VbaWindow", vbTextCompare) = 0 And InStr(1, cn, "Code", vbTextCompare) = 0 And InStr(1, cn, "wndclass", vbTextCompare) = 0 Then
        d("reason") = "focus_not_code_editor"
        Err.Clear
        Set WoPostF9ToFocus = d
        Exit Function
    End If
    WoPostMessageW h, WM_KEYDOWN, &H78, 1&
    WoPostMessageW h, WM_KEYUP, &H78, &HC0000001
    d("ok") = True
    d("reason") = "posted_f9"
    Err.Clear
    Set WoPostF9ToFocus = d
End Function

Public Function WoFrontHwnd() As Long
    WoFrontHwnd = WoGetForegroundWindow()
End Function

Public Function WoFrontRestore(ByVal h As Long) As Boolean
    ' Returns True only when SetForegroundWindow reported success - never claims silently.
    If h = 0 Then Exit Function
    WoFrontRestore = (WoSetForegroundWindow(h) <> 0)
End Function

' =====================================================================
' R2 + R12: the live desktop / input-channel probe. READ-ONLY.
'   Every value is MEASURED on every call from inside THIS process. Nothing is
'   inferred from the window-station NAME, and canSendInput is NOT derived from
'   GetForegroundWindow()==0: this add-in delivers input as PostMessage
'   (WM_KEYDOWN / WM_KEYUP / WM_MOUSEWHEEL) to windows of its OWN process, measured
'   to deliver while ANOTHER process owns the foreground. The handle OpenInputDesktop
'   opens is closed again before returning.
' =====================================================================
Private Function WoNz(ByVal s As String) As String
    Dim i As Long
    i = InStr(1, s, vbNullChar)
    If i = 0 Then WoNz = s Else WoNz = Left$(s, i - 1)
End Function

Private Function WoUserObjectNameOf(ByVal hObj As Long, ByRef outName As String, ByRef errNo As Long, ByRef errTx As String) As Boolean
    Dim buf As String
    Dim need As Long
    Dim rc As Long
    WoUserObjectNameOf = False
    outName = ""
    errNo = 0
    errTx = ""
    If hObj = 0 Then Exit Function
    On Error Resume Next
    Err.Clear
    buf = String$(512, vbNullChar)
    need = 0
    rc = WoGetUserObjectInformationW(hObj, WO_UOI_NAME, StrPtr(buf), 512, need)
    If rc = 0 And Err.Number = 0 And need > 0 And need <= 4096 Then
        buf = String$(need, vbNullChar)
        rc = WoGetUserObjectInformationW(hObj, WO_UOI_NAME, StrPtr(buf), need, need)
    End If
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Then Exit Function
    If rc = 0 Then Exit Function
    outName = WoNz(buf)
    WoUserObjectNameOf = True
End Function

' R2/R12: the input-desktop probe. FAIL-CLOSED: it reports what the API actually did, including
' the VB error when the call could not be made at all (a Declare missing its Alias clause raises
' VB error 453 here). A zero handle with NO error is a genuine live negative (no interactive
' input desktop), not a failure.
Private Function WoInputDesktopProbe(ByRef hRet As Long, ByRef errNo As Long, ByRef errTx As String) As Boolean
    WoInputDesktopProbe = False
    hRet = 0
    errNo = 0
    errTx = ""
    On Error Resume Next
    Err.Clear
    hRet = WoOpenInputDesktop(0, 0, WO_DESKTOP_READOBJECTS Or WO_DESKTOP_WRITEOBJECTS Or WO_DESKTOP_SWITCHDESKTOP)
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Then Exit Function
    If hRet <> 0 Then
        WoInputDesktopProbe = True
        Call WoCloseDesktop(hRet)
    End If
End Function

' Minimal JSON string escaping for raw API text carried in a refusal.
Private Function WoJsonQ(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, Chr$(34), "\" & Chr$(34))
    t = Replace(t, vbCrLf, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    WoJsonQ = Chr$(34) & t & Chr$(34)
End Function

' FAIL-CLOSED refusal for the desktop probe. No desktop/interactivity value is ever reported as a
' measurement unless its API actually answered; this shape is returned INSTEAD of a zero-filled
' picture that a caller could mistake for real values.
Private Function WoDesktopFailJson(ByVal d As Object, ByVal apiName As String, ByVal errNo As Long, ByVal errTx As String, ByVal returnedZero As Boolean) As String
    Dim why As String
    d("ok") = False
    d("failureKind") = "api_unavailable"
    d("failedApi") = apiName
    d("apiError") = errNo
    d("apiErrorText") = errTx
    d("measured") = False
    If returnedZero Then d("returnedZeroHandle") = True
    If errNo <> 0 Then
        why = "VB error " & CStr(errNo) & " (" & errTx & ")"
    Else
        why = "the API returned a null handle with no error"
    End If
    d("reason") = "the desktop probe could not obtain " & apiName & ": " & why
    d("note") = "FAIL-CLOSED: no desktop/interactivity value is reported unless its API answered, so this refusal is returned instead of a zero-filled picture. A Declare whose Alias clause is missing raises VB error 453 here."
    WoDesktopFailJson = JsonStringify(d)
End Function

' R12: the desktop{} fragment spliced into the input-family receipts. FAIL-CLOSED: when the
' input-desktop API cannot be called, the fragment says measured:false and carries the raw error
' instead of claiming a capability value.
Public Function WoDesktopFieldsJson() As String
    Dim hInput As Long
    Dim errNo As Long
    Dim errTx As String
    Dim interactive As Boolean
    Dim flagTxt As String
    interactive = WoInputDesktopProbe(hInput, errNo, errTx)
    flagTxt = "false"
    If interactive Then flagTxt = "true"
    If errNo <> 0 Then
        WoDesktopFieldsJson = Chr$(34) & "desktop" & Chr$(34) & ":{" & Chr$(34) & "interactive" & Chr$(34) & ":false," & Chr$(34) & "canSendInput" & Chr$(34) & ":false," & _
            Chr$(34) & "measured" & Chr$(34) & ":false," & Chr$(34) & "failureKind" & Chr$(34) & ":" & Chr$(34) & "api_unavailable" & Chr$(34) & "," & Chr$(34) & "apiError" & Chr$(34) & ":" & CStr(errNo) & "," & _
            Chr$(34) & "apiErrorText" & Chr$(34) & ":" & WoJsonQ(errTx) & "," & Chr$(34) & "note" & Chr$(34) & ":" & WoJsonQ("the input-desktop API could not be called, so no input-capability value is claimed and measured is false") & "}"
        Exit Function
    End If
    WoDesktopFieldsJson = Chr$(34) & "desktop" & Chr$(34) & ":{" & Chr$(34) & "interactive" & Chr$(34) & ":" & flagTxt & "," & Chr$(34) & "canSendInput" & Chr$(34) & ":" & flagTxt & "," & Chr$(34) & "measured" & Chr$(34) & ":true}"
End Function

' R12 / R2: the desktop_probe tool body.
' R12 / R2: the desktop_probe tool body. FAIL-CLOSED: each API is called with its own error trap
' and its result is checked; any failure returns ok:false + failureKind + the raw VB error instead
' of a zero-filled success. OpenInputDesktop returning 0 WITHOUT an error is a genuine live
' negative (no interactive input desktop) and is reported as such.
Public Function WoDesktopJson() As String
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    Dim hWinSta As Long, hDesk As Long, hInput As Long
    Dim winstaName As String, desktopName As String
    Dim errNo As Long, errTx As String
    Dim fg As Long, fgIsWin As Long
    Dim sid As Long, sidOk As Long
    Dim interactive As Boolean, fgReadable As Boolean
    Dim flagTxt As String
    d("tool") = "desktop_probe"
    d("probeBasis") = "GetProcessWindowStation + GetUserObjectInformationW(UOI_NAME) + GetThreadDesktop(GetCurrentThreadId) + OpenInputDesktop + GetForegroundWindow + IsWindow + ProcessIdToSessionId, each called with its own error trap; a value is only reported as measured when its API actually answered"
    On Error Resume Next
    Err.Clear
    hWinSta = WoGetProcessWindowStation()
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Or hWinSta = 0 Then
        WoDesktopJson = WoDesktopFailJson(d, "GetProcessWindowStation", errNo, errTx, (hWinSta = 0))
        Exit Function
    End If
    If Not WoUserObjectNameOf(hWinSta, winstaName, errNo, errTx) Then
        WoDesktopJson = WoDesktopFailJson(d, "GetUserObjectInformationW(UOI_NAME) on the window station", errNo, errTx, False)
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    hDesk = WoGetThreadDesktop(WoGetCurrentThreadId())
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Or hDesk = 0 Then
        WoDesktopJson = WoDesktopFailJson(d, "GetThreadDesktop(GetCurrentThreadId)", errNo, errTx, (hDesk = 0))
        Exit Function
    End If
    If Not WoUserObjectNameOf(hDesk, desktopName, errNo, errTx) Then
        WoDesktopJson = WoDesktopFailJson(d, "GetUserObjectInformationW(UOI_NAME) on the thread desktop", errNo, errTx, False)
        Exit Function
    End If
    sid = 0
    On Error Resume Next
    Err.Clear
    sidOk = WoProcessIdToSessionId(WoGetCurrentProcessId(), sid)
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Or sidOk = 0 Then
        WoDesktopJson = WoDesktopFailJson(d, "ProcessIdToSessionId(GetCurrentProcessId)", errNo, errTx, (sidOk = 0))
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    fg = WoGetForegroundWindow()
    errNo = Err.Number
    errTx = Err.Description
    Err.Clear
    On Error GoTo 0
    If errNo <> 0 Then
        WoDesktopJson = WoDesktopFailJson(d, "GetForegroundWindow", errNo, errTx, False)
        Exit Function
    End If
    fgIsWin = 0
    If fg <> 0 Then
        On Error Resume Next
        Err.Clear
        fgIsWin = WoIsWindow(fg)
        errNo = Err.Number
        errTx = Err.Description
        Err.Clear
        On Error GoTo 0
        If errNo <> 0 Then
            WoDesktopJson = WoDesktopFailJson(d, "IsWindow(foregroundHwnd)", errNo, errTx, False)
            Exit Function
        End If
    End If
    interactive = WoInputDesktopProbe(hInput, errNo, errTx)
    If errNo <> 0 Then
        WoDesktopJson = WoDesktopFailJson(d, "OpenInputDesktop", errNo, errTx, False)
        Exit Function
    End If
    fgReadable = (fg <> 0 And fgIsWin <> 0)
    flagTxt = "false"
    If interactive Then flagTxt = "true"
    d("ok") = True
    d("measured") = True
    d("windowStation") = winstaName
    d("windowStationHandle") = hWinSta
    d("desktop") = desktopName
    d("desktopHandle") = hDesk
    d("sessionId") = sid
    d("sessionIdOk") = (sidOk <> 0)
    d("interactive") = interactive
    d("canSendInput") = interactive
    d("inputDesktopOk") = interactive
    d("inputDesktopHandle") = hInput
    d("foregroundHwnd") = fg
    d("foregroundIsWindow") = (fgIsWin <> 0)
    d("foregroundReadable") = fgReadable
    d("note") = "every value above was returned by the API named in probeBasis and nothing was defaulted; this shape is only emitted after ALL of those APIs answered. interactive/canSendInput are LIVE values re-measured on every call and are NOT derived from the window-station name nor from GetForegroundWindow()==0. This add-in delivers input as PostMessage(WM_KEYDOWN/WM_KEYUP/WM_MOUSEWHEEL) to windows of its OWN process, measured to deliver while ANOTHER process owns the foreground. OpenInputDesktop returning 0 with no error is a genuine live negative and is shown by inputDesktopOk/inputDesktopHandle."
    WoDesktopJson = JsonStringify(d)
End Function

' =====================================================================
' PERF DIAGNOSTICS (no behaviour change, OFF unless VB6MCP_PERF_DIAG=1).
' One line per window names the per-tick period AND the microseconds spent in each
' segment, so 'who spends the idle CPU' is measured instead of argued.
' seg 0 = m_pipe.Tick (INCLUSIVE of seg 1), 1 = m_ctl.Tick, 2 = IngressPump,
' 3 = IdentityWatchTick (whole), 4 = the two VBIDE COM identity calls,
' 5 = WinOps.WheelTick, 6 = RegistrationTick.
' =====================================================================

Private Function WoQpcRaw(ByRef q As WO_QPC) As Double
    Dim hi As Double, lo As Double
    hi = q.hi
    If hi < 0 Then hi = hi + 4294967296#
    lo = q.lo
    If lo < 0 Then lo = lo + 4294967296#
    WoQpcRaw = hi * 4294967296# + lo
End Function

Public Function WoProfNow() As Double
    If Not m_pfOn Then Exit Function
    Dim q As WO_QPC
    If WoQueryPerformanceCounter(q) = 0 Then Exit Function
    WoProfNow = WoQpcRaw(q) * 1000000# / m_pfFreq
End Function

' R2 PERF GATE: the three tick handlers ask ONCE per pass and then branch on a LOCAL
' Boolean, so a diag-off tick costs ONE cross-module call instead of one per probe
' (measured ~19). Diag-on output is unchanged: same segment numbers, same order.
Public Function WoProfOn() As Boolean
    WoProfOn = m_pfOn
End Function

Public Sub WoProfInit(ByVal tickMs As Long)
    On Error Resume Next
    m_pfTickMs = tickMs
    m_pfOn = (Trim$(Environ$("VB6MCP_PERF_DIAG")) = "1")
    If Not m_pfOn Then Exit Sub
    Dim f As WO_QPC
    If WoQueryPerformanceFrequency(f) = 0 Then
        m_pfOn = False
        Exit Sub
    End If
    m_pfFreq = WoQpcRaw(f)
    If m_pfFreq <= 0 Then
        m_pfOn = False
        Exit Sub
    End If
    m_pfWindowAt = WoProfNow()
    m_pfCpuStart = WoThreadCpuUs()
    Err.Clear
End Sub

Public Sub WoProfTickCount()
    If Not m_pfOn Then Exit Sub
    m_pfTicks = m_pfTicks + 1
End Sub

' patch O: which method the last served RPC was, so a slow window is attributable
' to a NAME instead of to the pipe. Truncated: the perf line must not grow per request.
Public Sub WoProfMethod(ByVal m As String)
    If Len(m) > 40 Then m = Left$(m, 40)
    m = Replace(m, " ", "_")
    m = Replace(m, Chr$(34), "_")
    m_pfLastMethod = m
    Dim key As String
    Dim nm As String
    key = LCase$(m)
    nm = key
    If Left$(key, 11) = "tools/call:" Then nm = Mid$(key, 12)
    If key = "tools/list" Then
        m_rpcToolsListN = m_rpcToolsListN + 1
        m_rpcBucket = 2
    ElseIf nm = "mcp_info" Then
        m_rpcMcpInfoN = m_rpcMcpInfoN + 1
        m_rpcBucket = 1
    ElseIf nm = "get_project_info" Then
        m_rpcOtherN = m_rpcOtherN + 1
        m_rpcBucket = 5
    ElseIf Left$(nm, 4) = "ext_" Then
        m_rpcExtN = m_rpcExtN + 1
        m_rpcBucket = 4
    Else
        m_rpcOtherN = m_rpcOtherN + 1
        m_rpcBucket = 3
    End If
End Sub

Public Sub WoProfMethodDone(ByVal us As Double)
    If m_rpcBucket = 1 Then
        m_rpcMcpInfoUs = m_rpcMcpInfoUs + us
    ElseIf m_rpcBucket = 2 Then
        m_rpcToolsListUs = m_rpcToolsListUs + us
    ElseIf m_rpcBucket = 4 Then
        m_rpcExtUs = m_rpcExtUs + us
    ElseIf m_rpcBucket = 5 Then
        m_rpcGetProjectInfoUs = m_rpcGetProjectInfoUs + us
    Else
        m_rpcOtherUs = m_rpcOtherUs + us
    End If
    m_rpcBucket = 0
End Sub

Public Sub WoProfNestedTick(ByVal us As Double)
    m_nestedRpcCount = m_nestedRpcCount + 1
    m_nestedRpcTickUs = m_nestedRpcTickUs + us
End Sub

' patch O tools/list cache counters (cumulative; the hub can subtract two windows).
Public Sub WoProfCacheHit(ByVal bytes As Long)
    m_pfCacheHits = m_pfCacheHits + 1
    m_pfCacheBytes = bytes
End Sub

Public Sub WoProfCacheMiss(ByVal bytes As Long)
    m_pfCacheMisses = m_pfCacheMisses + 1
    m_pfCacheBytes = bytes
End Sub

' patch O v2: one frame was BUILT (payload + 4-byte header) and handed to the
' write lane. bytes is the wire size of the whole frame.
Public Sub WoProfOutFrame(ByVal bytes As Long)
    m_pfOutFrames = m_pfOutFrames + 1
    m_pfOutBytes = m_pfOutBytes + bytes
End Sub

' patch O v2: a write that completed SYNCHRONOUSLY (NpWriteFile returned ok<>0).
' Before this counter existed those bytes were never added to anything.
Public Sub WoProfWriteSync(ByVal bytes As Long)
    m_pfWriteSyncBytes = m_pfWriteSyncBytes + bytes
End Sub

' patch O v2: bytes reaped by the overlapped completion path (state 4).
Public Sub WoProfOutReap(ByVal bytes As Long)
    m_pfOutReapBytes = m_pfOutReapBytes + bytes
End Sub

' patch O v2: size of the LAST request frame and the response it produced.
Public Sub WoProfReqResp(ByVal reqBytes As Long, ByVal respBytes As Long)
    m_pfLastReqBytes = reqBytes
    m_pfLastRespBytes = respBytes
End Sub

Public Sub WoProfAdd(ByVal seg As Long, ByVal us As Double)
    If Not m_pfOn Then Exit Sub
    If seg < 0 Or seg >= WO_PF_SEGS Then Exit Sub
    If us <= 0 Then
        m_pfN(seg) = m_pfN(seg) + 1
        Exit Sub
    End If
    m_pfUs(seg) = m_pfUs(seg) + us
    m_pfN(seg) = m_pfN(seg) + 1
End Sub

' Returns ONE line when the window is due, else "". The caller logs it, so this
' module stays free of the host's logging policy.
Public Function WoProfConsume(ByVal pumpCalls As Long) As String
    WoProfConsume = ""
    If Not m_pfOn Then Exit Function
    On Error Resume Next
    Dim nowUs As Double, elapsed As Double, cpuNow As Double, cpuUs As Double
    nowUs = WoProfNow()
    elapsed = nowUs - m_pfWindowAt
    If elapsed < WO_PF_WINDOW_US Then Exit Function
    cpuNow = WoThreadCpuUs()
    cpuUs = cpuNow - m_pfCpuStart
    If cpuUs < 0 Then cpuUs = 0
    Dim s As String
    s = "perf diag tickMs=" & CStr(m_pfTickMs)
    s = s & " ticks=" & CStr(m_pfTicks)
    s = s & " elapsedUs=" & CStr(CLng(elapsed))
    s = s & " cpuUs=" & CStr(CLng(cpuUs))
    s = s & " cpuPct=" & CStr(CLng(cpuUs * 100# / elapsed))
    s = s & " qpcFreq=" & CStr(CLng(m_pfFreq))
    s = s & " pumpCalls=" & CStr(pumpCalls)
    s = s & " segPipeUs=" & CStr(CLng(m_pfUs(0))) & " pipeCount=" & CStr(m_pfN(0))
    s = s & " segCtlUs=" & CStr(CLng(m_pfUs(1))) & " ctlCount=" & CStr(m_pfN(1))
    s = s & " segIngressUs=" & CStr(CLng(m_pfUs(2))) & " ingressCount=" & CStr(m_pfN(2))
    s = s & " segIdentityUs=" & CStr(CLng(m_pfUs(3))) & " identityCount=" & CStr(m_pfN(3))
    s = s & " identComUs=" & CStr(CLng(m_pfUs(4))) & " identComCalls=" & CStr(m_pfN(4))
    s = s & " segWheelUs=" & CStr(CLng(m_pfUs(5))) & " wheelCount=" & CStr(m_pfN(5))
    s = s & " segRegUs=" & CStr(CLng(m_pfUs(6))) & " regCount=" & CStr(m_pfN(6))
    s = s & " segEnsurePipeUs=" & CStr(CLng(m_pfUs(7))) & " ensurePipeCount=" & CStr(m_pfN(7))
    s = s & " segPollConnectUs=" & CStr(CLng(m_pfUs(8))) & " pollConnectCount=" & CStr(m_pfN(8))
    s = s & " segServiceUs=" & CStr(CLng(m_pfUs(9))) & " serviceCount=" & CStr(m_pfN(9))
    s = s & " segPollReadUs=" & CStr(CLng(m_pfUs(10))) & " pollReadCount=" & CStr(m_pfN(10))
    s = s & " segPollWriteUs=" & CStr(CLng(m_pfUs(11))) & " pollWriteCount=" & CStr(m_pfN(11))
    s = s & " segTickLoopUs=" & CStr(CLng(m_pfUs(12))) & " tickLoopCount=" & CStr(m_pfN(12))
    s = s & " segPumpUs=" & CStr(CLng(m_pfUs(13))) & " pumpDirectCount=" & CStr(m_pfN(13))
    s = s & " segRpcUs=" & CStr(CLng(m_pfUs(14))) & " rpcCount=" & CStr(m_pfN(14))
    s = s & " lastMethod=" & m_pfLastMethod
    s = s & " segTimerUs=" & CStr(CLng(m_pfUs(15))) & " timerCount=" & CStr(m_pfN(15))
    s = s & " segSvcPeekUs=" & CStr(CLng(m_pfUs(16))) & " svcPeekCount=" & CStr(m_pfN(16))
    s = s & " toolsCacheHits=" & CStr(m_pfCacheHits) & " toolsCacheMisses=" & CStr(m_pfCacheMisses)
    s = s & " toolsCacheBytes=" & CStr(m_pfCacheBytes)
    s = s & " segMcpShaUs=" & CStr(CLng(m_pfUs(19))) & " mcpShaCount=" & CStr(m_pfN(19))
    s = s & " segMcpEnvUs=" & CStr(CLng(m_pfUs(20))) & " mcpEnvCount=" & CStr(m_pfN(20))
    s = s & " segMcpJsonUs=" & CStr(CLng(m_pfUs(21))) & " mcpJsonCount=" & CStr(m_pfN(21))
    s = s & " rpcToolsListN=" & CStr(m_rpcToolsListN) & " rpcMcpInfoN=" & CStr(m_rpcMcpInfoN)
    s = s & " rpcOtherN=" & CStr(m_rpcOtherN) & " rpcExtN=" & CStr(m_rpcExtN)
    s = s & " segWheelFocusApiUs=" & CStr(CLng(m_pfUs(23))) & " wheelFocusApiCount=" & CStr(m_pfN(23))
    s = s & " segWheelAttachUs=" & CStr(CLng(m_pfUs(25))) & " wheelAttachCount=" & CStr(m_pfN(25))
    s = s & " rpcMcpInfoUsLife=" & CStr(CLng(m_rpcMcpInfoUs)) & " rpcToolsListUsLife=" & CStr(CLng(m_rpcToolsListUs))
    s = s & " rpcOtherUsLife=" & CStr(CLng(m_rpcOtherUs)) & " rpcExtUsLife=" & CStr(CLng(m_rpcExtUs))
    s = s & " rpcGetProjectInfoUsLife=" & CStr(CLng(m_rpcGetProjectInfoUs))
    s = s & " nestedRpcCount=" & CStr(m_nestedRpcCount) & " nestedRpcTickUs=" & CStr(CLng(m_nestedRpcTickUs))
    s = s & " outFrames=" & CStr(m_pfOutFrames) & " outBytes=" & CStr(m_pfOutBytes)
    s = s & " writeBytesSync=" & CStr(m_pfWriteSyncBytes) & " outReapBytes=" & CStr(m_pfOutReapBytes)
    s = s & " lastReqBytes=" & CStr(m_pfLastReqBytes) & " lastRespBytes=" & CStr(m_pfLastRespBytes)
    s = s & " segWheelDrainUs=" & CStr(CLng(m_pfUs(17))) & " wheelDrainCount=" & CStr(m_pfN(17))
    s = s & " segWheelFocusUs=" & CStr(CLng(m_pfUs(18))) & " wheelFocusCount=" & CStr(m_pfN(18))
    ' H1: if the pipe cost is paid per unit of TIME instead of per tick, two rounds with
    ' different tick periods are only comparable normalised per SECOND. Report both, plus
    ' the MEASURED period so nobody has to assume tickMs is what actually ran.
    If m_pfTicks > 0 Then
        s = s & " tickPeriodUs=" & CStr(CLng(elapsed / m_pfTicks))
    End If
    Dim per As Double
    If elapsed > 0 Then per = 1000# / elapsed
    s = s & " pipeMsPerSec=" & CStr(CLng(m_pfUs(0) * per))
    s = s & " pipeSubMsPerSec=ensure" & CStr(CLng(m_pfUs(7) * per))
    s = s & "/connect" & CStr(CLng(m_pfUs(8) * per))
    s = s & "/service" & CStr(CLng(m_pfUs(9) * per))
    s = s & "/read" & CStr(CLng(m_pfUs(10) * per))
    s = s & "/write" & CStr(CLng(m_pfUs(11) * per))
    s = s & "/loop" & CStr(CLng(m_pfUs(12) * per))
    s = s & " wheelMsPerSec=" & CStr(CLng(m_pfUs(5) * per))
    Dim i As Long
    For i = 0 To WO_PF_SEGS - 1
        m_pfUs(i) = 0#
        m_pfN(i) = 0
    Next
    m_pfTicks = 0
    m_pfWindowAt = nowUs
    m_pfCpuStart = cpuNow
    WoProfConsume = s
    Err.Clear
End Function

' =====================================================================
' THREAD CPU (wall vs CPU separation) + PUMP segment + environment{}.
' =====================================================================

Private Function WoFtRaw(ByRef f As WO_FILETIME) As Double
    Dim hi As Double, lo As Double
    hi = f.hi
    If hi < 0 Then hi = hi + 4294967296#
    lo = f.lo
    If lo < 0 Then lo = lo + 4294967296#
    WoFtRaw = hi * 4294967296# + lo
End Function

' CPU microseconds consumed by THIS thread (kernel + user). QPC wall time cannot
' tell a blocking wait from real work; this can.
Public Function WoThreadCpuUs() As Double
    On Error Resume Next
    Dim c As WO_FILETIME, e As WO_FILETIME, k As WO_FILETIME, u As WO_FILETIME
    If WoGetThreadTimes(WoGetCurrentThread(), c, e, k, u) = 0 Then Exit Function
    WoThreadCpuUs = (WoFtRaw(k) + WoFtRaw(u)) / 10#
    Err.Clear
End Function

' Record one DIRECT DoEvents span (segment 13). Only the call sites that actually
' call DoEvents feed this, so the number means what it says.
Public Sub WoPumpAdd(ByVal us As Double)
    If Not m_pfOn Then Exit Sub
    If us < 0 Then us = 0
    m_pfUs(13) = m_pfUs(13) + us
    m_pfN(13) = m_pfN(13) + 1
End Sub

Private Function WoHostExeVersion() As String
    On Error Resume Next
    WoHostExeVersion = "unknown"
    Dim buf As String, n As Long, dummy As Long, sz As Long
    buf = String$(1024, 0)
    n = WoGetModuleFileNameW(0&, StrPtr(buf), 512&)
    If n <= 0 Then Exit Function
    Dim pathW As String
    pathW = Left$(buf, n)
    sz = WoGetFileVersionInfoSizeW(StrPtr(pathW), dummy)
    If sz <= 0 Then Exit Function
    Dim data As String
    data = String$(sz, 0)
    If WoGetFileVersionInfoW(StrPtr(pathW), 0&, sz, StrPtr(data)) = 0 Then Exit Function
    Dim langs As Variant, lang As Variant, subBlk As String
    Dim pv As Long, pu As Long, s As String, m As Long
    langs = Array("040904b0", "080404b0", "000004b0", "040704b0")
    For Each lang In langs
        subBlk = "\StringFileInfo\" & CStr(lang) & "\FileVersion"
        pv = 0
        pu = 0
        If WoVerQueryValueW(StrPtr(data), StrPtr(subBlk), pv, pu) <> 0 Then
            If pv <> 0 And pu > 0 Then
                m = WoLstrlenW(pv)
                If m > 0 And m < 128 Then
                    s = String$(m + 1, 0)
                    WoLStrCpyW StrPtr(s), pv
                    WoHostExeVersion = Left$(s, m)
                    Err.Clear
                    Exit Function
                End If
            End If
        End If
    Next
    Err.Clear
End Function

' environment block for mcp_info (cross-machine fingerprint). Computed ONLY when
' mcp_info is called - it is never touched by a timer.
' LANGID -> BCP-47 tag. "und" when the language name cannot be read: an unknown tag is
' honest, a fabricated one is not.
Private Function WoBcp47FromLangId(ByVal langid As Long) As String
    ' One call, whole tag: LOCALE_SNAME already carries language + optional script + region,
    ' so there is no stitching to get wrong and no case to guess. MAKELCID(langid,
    ' SORT_DEFAULT) is the langid itself (sort 0), hence the mask.
    On Error Resume Next
    WoBcp47FromLangId = "und"
    Dim buf As String, n As Long, s As String
    buf = String$(85, 0)
    n = WoGetLocaleInfoW(langid And &HFFFF&, WO_LOCALE_SNAME, StrPtr(buf), 85&)
    If n > 1 Then
        s = Left$(buf, n - 1)
        If Len(s) > 0 Then WoBcp47FromLangId = s
    End If
    Err.Clear
End Function

Public Function EnvironmentInfoObj() As Object
    ' CONTRACT SHAPE (asserted by tests/matrix/suites/fingerprint.mjs): six keys, FIXED types -
    ' schema/vb6Version/uiLang = String, ansiCodePage/osBuild = Long (NUMBERS, not text),
    ' interactive = Boolean. No auxiliary keys: an extra contract member nobody asked for is
    ' worse than a missing one.
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    d("schema") = ENV_SCHEMA_ID
    d("vb6Version") = WoHostExeVersion()
    d("ansiCodePage") = WoGetACP()
    ' uiLang is a BCP-47 TAG ("zh-CN"): the suite matches ^[a-z]{2,3}(-[A-Za-z0-9]{2,8})*$.
    d("uiLang") = WoBcp47FromLangId(WoGetUserDefaultUILanguage())
    ' osBuild is a NUMBER (integer >= 7600), so the registry string is converted.
    Dim b As String, nb As Long
    b = RegGetString("SOFTWARE\Microsoft\Windows NT\CurrentVersion", "CurrentBuildNumber")
    nb = 0
    Err.Clear
    If Len(Trim$(b)) > 0 Then nb = CLng(Trim$(b))
    Err.Clear
    d("osBuild") = nb
    ' Same probe desktop_probe uses (OpenInputDesktop, handle closed again). FALSE is a
    ' MEASURED negative and is reported as such - never papered over with a default.
    Dim hInput As Long, errNo As Long, errTx As String
    d("interactive") = WoInputDesktopProbe(hInput, errNo, errTx)
    Set EnvironmentInfoObj = d
    Err.Clear
End Function

Public Sub WoProfFindCmd()
    m_findCmdCount = m_findCmdCount + 1
End Sub
' =======================================================================
' HOST TICK BRIDGE (Win32 TIMERPROC) - the DEFAULT (ServerForm) host.
'
' WHY THIS EXISTS. The default host used to drive the pipe tick from a VB6 Timer
' CONTROL. A Timer control event does NOT re-enter: while the previous event is
' still on the stack the runtime refuses to raise it again, so a DoEvents issued
' from inside a long tool call (Util.MaybePump) could not dispatch another tick.
' MEASURED consequence: nestedTicks stayed 0 and PipeTransport's progress frame
' was never produced under that host, while the Win32-timer host measured
' nestedTicks=710 / progressFrames=5 on the same code.
'
' A SetTimer TIMERPROC is called by the OS from the thread's message pump, so the
' SAME DoEvents DOES dispatch it; that nested pass reaches PipeTransport.Tick,
' which recognises the re-entry and emits the progress frame WITHOUT re-entering
' request handling. Same mechanism the optional code-pane wheel already uses
' (WoSetTimer + WheelTimerProc) - reused, not re-invented.
'
' hWnd = 0 creates a CALLBACK timer: no window, and its id is allocated by the
' system, so it cannot collide with any window timer (WinsockApi keeps its own
' m_hWnd + id 1).
' =======================================================================
Public Function WoHostTickRunning() As Boolean
    WoHostTickRunning = (m_hostTickId <> 0)
End Function

Public Function WoHostTickPeriodMs() As Long
    WoHostTickPeriodMs = m_hostTickMs
End Function

' Attach the sink that receives Public Sub HostTick(). Late-bound on purpose:
' WinOps is a standard module and must not name ServerForm at compile time.
Public Function WoHostTickStart(ByVal sink As Object, ByVal ms As Long) As Boolean
    On Error GoTo Failed
    WoHostTickStop
    If ms < 1 Then ms = 1
    If ms > 60000 Then ms = 60000
    Set m_hostTickSink = sink
    m_hostTickMs = ms
    m_hostTickId = WoSetTimer(0, 0, ms, AddressOf HostTickProc)
    WoHostTickStart = (m_hostTickId <> 0)
    Exit Function
Failed:
    m_hostTickId = 0
    WoHostTickStart = False
End Function

Public Sub WoHostTickStop()
    On Error Resume Next
    If m_hostTickId <> 0 Then WoKillTimer 0, m_hostTickId
    m_hostTickId = 0
End Sub

' Change the period of a RUNNING tick. A callback timer owns no reusable id, so
' this is kill+set; the old id is killed first, so at most one timer is ever alive.
Public Function WoHostTickSetPeriod(ByVal ms As Long) As Boolean
    If m_hostTickSink Is Nothing Then Exit Function
    WoHostTickSetPeriod = WoHostTickStart(m_hostTickSink, ms)
End Function

Public Sub WoHostTickDetach()
    WoHostTickStop
    Set m_hostTickSink = Nothing
    m_hostTickMs = 0
End Sub

' The TIMERPROC. Same prototype as WheelTimerProc because it IS a TIMERPROC.
' RE-ENTRANCY DISCIPLINE (the callback must never recurse). Exactly TWO levels are
' allowed: level 1 is the full watchdog pass, level 2 is the nested pass a DoEvents
' inside a long tool call dispatches, where the sink runs ONLY the pipe pass (that
' is what becomes a progress frame). Anything deeper is refused outright, so the
' callback cannot fan out.
Public Sub HostTickProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    On Error Resume Next
    If m_hostTickSink Is Nothing Then Exit Sub
    If m_hostTickDepth >= 2 Then
        Err.Clear
        Exit Sub
    End If
    m_hostTickDepth = m_hostTickDepth + 1
    On Error GoTo Done
    m_hostTickSink.HostTick
Done:
    m_hostTickDepth = m_hostTickDepth - 1
    If m_hostTickDepth < 0 Then m_hostTickDepth = 0
    Err.Clear
End Sub
