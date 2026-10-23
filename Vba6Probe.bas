Attribute VB_Name = "Vba6Probe"
Option Explicit

' VBA6.DLL debug-engine exports (undocumented, internal). IDE-ONLY.
' Verified against the real export table by tools/check_vba6_declares.py.
' Signatures follow 内部逆向记录（VBA6 答疑） + 内部逆向复核记录
' (note: CreateIExprSrvObj is 3 args, NOT the bogus ret 1064).
Private Declare Function EbMode Lib "VBA6.DLL" () As Long
Private Declare Function EbGetStepMode Lib "VBA6.DLL" () As Long
Private Declare Function EbSetStepMode Lib "VBA6.DLL" (ByVal stepMode As Long) As Long
Private Declare Function EbSetStepToCursor Lib "VBA6.DLL" (ByVal enable As Long) As Long
Private Declare Function EbFExitBreakMode Lib "VBA6.DLL" () As Long
Private Declare Function EbGetCallstackCount Lib "VBA6.DLL" (ByRef pCount As Long) As Long
' arg5 is a POINTER (out long*), not a value: the generated code does
'   mov eax,[ebp+0x18] ; mov dword ptr [eax],1
' so passing ByVal 0 wrote to address 0 and raised 0xc0000005. Verified by
' disassembling VBA6.DLL RVA 0x13348D..0x133492 against the live crash.
' arg2 = module/display name (BSTR*), arg3 = PROCEDURE name (BSTR*),
' arg4 = 32-bit value written through a pointer, arg5 = success flag (1/0).
Private Declare Function EbGetCallstackFunction Lib "VBA6.DLL" (ByVal frameIndex As Long, ByRef modName As String, ByRef procName As String, ByRef a4 As Long, ByRef a5 As Long) As Long
Private Declare Function EbGetCurrentFunction Lib "VBA6.DLL" (ByRef pOut As Long, ByRef pOut2 As Long) As Long
Private Declare Function EbGetDebWndCallstackIndex Lib "VBA6.DLL" (ByRef pIndex As Long) As Long
Private Declare Function EbIsInEvalWatch Lib "VBA6.DLL" () As Long
' R7 solved: EbIsValidIdent is a TWO-argument function --
' (LPCWSTR, LONG* out). VB6.EXE's own call sites push &local then the wide string
' and read the out slot afterwards. Passing one argument made it write 1 through a
' garbage pointer ([ebp+0Ch]) -> 0xc0000005 at VBA6.DLL+0x12719C.
Private Declare Function EbIsValidIdent Lib "VBA6.DLL" (ByVal pwszName As Long, ByRef pResult As Long) As Long
Private Declare Function SysStringLen Lib "oleaut32" (ByVal pbstr As Long) As Long
' ---------------------------------------------------------------------------
' EbExecuteLine -- the Immediate-window line executor.
' A recheck of EbExecuteLine overturned the earlier
' verdict: the WideCharToMultiByte call at 0x0FBB70AC is only the conversion
' prologue; the function then loads the engine singleton and calls the real
' executor sub_FB7317A (RVA 0xE317A) = compile + evaluate + print pipeline.
' Independently verified here against the real export table:
'   VBA6.DLL EbExecuteLine  ordinal 147  RVA 0x127065   4 args, retn 10h
' Because the function loads ECX itself (mov ecx,[0FC0D15C]) a plain Declare
' needs NO asm trampoline -- but that global must be non-zero, so the probe
' refuses to call when it reads 0. Argument semantics are still candidates.
Private Declare Function EbExecuteLine Lib "VBA6.DLL" (ByVal pwszLine As Long, ByVal a2 As Long, ByVal a3 As Long, ByVal a4 As Long) As Long
' An earlier batch: read the engine global through the real module base (never hardcode the VA).
Private Declare Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As Long) As Long
' Note: the BSTR form must use a FRESH, independently allocated BSTR.
' Passing a VB string's own buffer (StrPtr) made the worker free it, which then
' raised "subscript out of range" in the caller. The copy is deliberately never
' freed here: the worker may already have freed it (double free would crash).
Private Declare Function SysAllocString Lib "oleaut32" (ByVal pwsz As Long) As Long
' VB6 cannot call an arbitrary address with 2 arguments, so build a tiny
' __stdcall thunk in executable memory and call it through CallWindowProcA:
'     push a2 ; push a1 ; mov eax,addr ; call eax ; ret 16
' The callee pops its own 8 bytes (ret 8); the closing "ret 16" removes the four
' arguments CallWindowProcA forwarded, so the caller's stack stays balanced.
Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function CallWindowProcA Lib "user32" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dst As Any, ByVal src As Long, ByVal cb As Long)
' write TO a raw address (the eval-scope word probe needs this direction)
Private Declare Sub CopyMemoryToPtr Lib "kernel32" Alias "RtlMoveMemory" (ByVal dst As Long, ByRef src As Any, ByVal cb As Long)

' ---------------------------------------------------------------------------
' Thread identity.
' The claim under test is that every probe that enters the VB6 runtime crashes
' because it runs OFF the VB6 main thread (STA/TLS). That claim has to be
' MEASURED before any main-thread dispatch machinery is built, so the standard
' probe response now always reports the calling thread and the thread that owns
' the IDE's main window.
'   GetWindowThreadProcessId RETURNS the thread id (its 2nd arg is the PID).
' ---------------------------------------------------------------------------
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
' Readability guard. mem_dump's numeric range check is NOT enough: reading an
' unmapped address (measured with 0x00100000) takes the whole IDE down.
Private Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal cb As Long) As Long
Private Declare Function FindWindowW Lib "user32" (ByVal lpClassName As Long, ByVal lpWindowName As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, ByRef lpdwProcessId As Long) As Long

' ---------------------------------------------------------------------------
' CALLER-SIDE addressing (VB6.EXE).
' The Add-In runs INSIDE VB6.EXE, so VB6.EXE is the host image and its
' internals are reachable with the same thunk used for VBA6.DLL. Legacy
' VB6.EXE / VBA6.DLL are linked WITHOUT /DYNAMICBASE, so if hostBase reads
' 0x00400000 an address from an IDA listing of the EXE is ALREADY a valid
' absolute address (VA = RVA + 0x400000) and needs no conversion at all.
' Everything below is READ-ONLY except the existing rva_call thunk.
' ---------------------------------------------------------------------------
Private Declare Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As Long, ByVal lpFileName As Long, ByVal nSize As Long) As Long
' Resolve an export BY NAME, so no probe depends on a hardcoded RVA.
Private Declare Function GetProcAddress Lib "kernel32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
' The engine allocates the returned BSTRs, the CALLER frees them.
Private Declare Sub SysFreeString Lib "oleaut32" (ByVal pbstr As Long)

' In-process window enumeration. VB6 requires these
' Declare statements to live in the module DECLARATION section, i.e. before the
' first procedure -- inserting them next to the code that uses them is a
' compile error ("only comments may appear after End Function").
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal uCmd As Long) As Long
Private Declare Function GetClassNameW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowTextW Lib "user32" (ByVal hWnd As Long, ByVal lpString As Long, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function GetClassLongW Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Function HostBase() As Long
    HostBase = GetModuleHandleW(0)
End Function

Public Function Vba6Base() As Long
    Vba6Base = GetModuleHandleW(StrPtr("VBA6.DLL"))
End Function

' PE SizeOfImage straight out of the mapped headers (no psapi dependency).
Public Function PeImageSize(ByVal hMod As Long) As Long
    Dim e_lfanew As Long, so As Long
    e_lfanew = 0
    so = 0
    If hMod = 0 Then Exit Function
    If IsBadReadPtr(hMod, 64) <> 0 Then Exit Function
    CopyMemory e_lfanew, hMod + &H3C, 4
    If e_lfanew <= 0 Or e_lfanew > &H1000 Then Exit Function
    If IsBadReadPtr(hMod + e_lfanew, 64) <> 0 Then Exit Function
    CopyMemory so, hMod + e_lfanew + &H50, 4
    PeImageSize = so
End Function

Private Function ModPath(ByVal hMod As Long) As String
    Dim buf As String, n As Long
    On Error Resume Next
    buf = String$(520, Chr$(0))
    n = GetModuleFileNameW(hMod, StrPtr(buf), 520)
    If n > 0 Then ModPath = Left$(buf, n)
    Err.Clear
End Function

' Reported by every probe: which images are mapped, where, and whether an IDA
' RVA can be used verbatim as an absolute address.
Public Function ModuleInfoJson() As String
    Dim hb As Long, vm As Long, hs As Long, vs As Long
    Dim s As String
    On Error Resume Next
    hb = HostBase()
    vm = Vba6Base()
    hs = PeImageSize(hb)
    vs = PeImageSize(vm)
    s = Q("hostBase") & ":" & CStr(hb)
    s = s & "," & Q("hostBaseHex") & ":" & Q("0x" & Hex$(hb))
    s = s & "," & Q("hostImageSize") & ":" & CStr(hs)
    s = s & "," & Q("hostPath") & ":" & Q(ModPath(hb))
    s = s & "," & Q("vba6Base") & ":" & CStr(vm)
    s = s & "," & Q("vba6BaseHex") & ":" & Q("0x" & Hex$(vm))
    s = s & "," & Q("vba6ImageSize") & ":" & CStr(vs)
    s = s & "," & Q("vba6Path") & ":" & Q(ModPath(vm))
    ' true => an address in an IDA listing of the EXE can be used verbatim
    s = s & "," & Q("hostVaEqIdaVa") & ":" & B(hb = &H400000)
    ModuleInfoJson = s
    Err.Clear
End Function

' Accepts "vb6+0x1234" / "vba6+0x1234" / "0x1234". A PLAIN rva keeps its
' historic meaning (VBA6.DLL), so no existing caller changes behaviour.
Public Function ResolveRvaSpec(ByVal spec As String, ByRef whichOut As String) As Long
    Dim t As String, lc As String
    t = Trim$(spec)
    whichOut = "vba6"
    lc = LCase$(t)
    If Len(lc) >= 5 Then
        If Left$(lc, 5) = "vba6+" Then
            t = Mid$(t, 6)
        ElseIf Left$(lc, 4) = "vb6+" Then
            whichOut = "vb6"
            t = Mid$(t, 5)
        ElseIf Left$(lc, 5) = "hexe+" Then
            whichOut = "vb6"
            t = Mid$(t, 6)
        End If
    ElseIf Len(lc) = 4 Then
        If Left$(lc, 4) = "vb6+" Then
            whichOut = "vb6"
            t = ""
        End If
    End If
    ResolveRvaSpec = ParseArgToken(t)
End Function

Private Function Q(ByVal v As String) As String
    Dim t As String
    t = Replace(v, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    Q = Chr$(34) & t & Chr$(34)
End Function

Private Function B(ByVal v As Boolean) As String
    If v Then B = "true" Else B = "false"
End Function
' R13: the uniform probe-failure envelope. ok=false + failureKind + reason are added to
' FAILURE paths only; every successful probe answer stays byte-identical to before.
Private Function FailErr(ByVal s As String, ByVal kind As String, ByVal reason As String) As String
    FailErr = s & "," & Q("ok") & ":false," & Q("failureKind") & ":" & Q(kind) & "," & Q("reason") & ":" & Q(reason) & "}"
End Function

Private Function FailSkip(ByVal s As String, ByVal kind As String, ByVal reason As String) As String
    FailSkip = s & "," & Q("ok") & ":false," & Q("failureKind") & ":" & Q(kind) & "," & Q("reason") & ":" & Q(reason) & "}"
End Function

Private Sub AddEnt(ByRef s As String, ByVal id As String, ByVal fn As String, ByVal arity As Long, ByVal pre As String, ByVal ro As Boolean, ByVal note As String)
    If Len(s) > 1 Then s = s & ","
    s = s & "{" & Chr$(34) & "id" & Chr$(34) & ":" & Q(id) & "," & Chr$(34) & "fn" & Chr$(34) & ":" & Q(fn) & "," & Chr$(34) & "arity" & Chr$(34) & ":" & CStr(arity) & "," & Chr$(34) & "precondition" & Chr$(34) & ":" & Q(pre) & "," & Chr$(34) & "readOnly" & Chr$(34) & ":" & B(ro) & "," & Chr$(34) & "note" & Chr$(34) & ":" & Q(note) & "}"
End Sub

' Ordered probe plan: read-only first, writes last. Run only inside the IDE process.
Public Function ProbePlanJson() As String
    Dim s As String
    s = "["
    AddEnt s, "P1", "EbMode", 0, "any", True, "Use GetIdeState instead; EbMode only returns 0 or 2."
    AddEnt s, "P2", "EbGetStepMode", 0, "any", True, "Step mode read."
    AddEnt s, "P3", "EbIsInEvalWatch", 0, "any", True, "Bit flag; implementation is mov eax,[g]; and eax,100h; ret."
    AddEnt s, "P4", "EbGetCallstackCount", 1, "break", True, "Out param = frame count."
    AddEnt s, "P5", "EbGetCallstackFunction", 5, "break", True, "arg2 is BSTR* -> ByRef String; a3/a4 semantics unresolved."
    AddEnt s, "P6", "EbGetCurrentFunction", 2, "break", True, "Two out params."
    AddEnt s, "P7", "EbGetDebWndCallstackIndex", 1, "break", True, "Call Stack window selected frame."
    AddEnt s, "P8", "EbFExitBreakMode", 0, "break", True, "Status query, not an action."
    AddEnt s, "P9", "EbIsValidIdent", 1, "design", True, "Identifier validity (useful for lint)."
    AddEnt s, "P10", "EbSetStepToCursor", 1, "break", False, "WRITE; uses current editor selection."
    AddEnt s, "P11", "EbSetStepMode", 1, "break", False, "WRITE; step mode."
    s = s & "]"
    ProbePlanJson = s
End Function

' =====================================================================
' ACTUAL EXECUTION (IDE process only)
' =====================================================================
' The plan above is inert; this is the entry point that really calls the
' debug engine. Order matters: EbMode first as a safety gate, read-only
' probes only, writes then gated behind force.
'
' EbMode returns 0 (not in a debug context) or 2 (break context) - the
' RE notes are explicit that it never returns anything else. We use
' <> 0 as the "safe to touch the callstack" gate and report the raw
' value so the assumption stays visible.

' The engine hands back a UTF-16 buffer while the BSTR length ends up as a BYTE
' count, so every second character is a NUL ("d\0b\0g\0"). Keep that as a
' heuristic until RE pins down the real arg2 type.
' Read a BSTR that lives at an absolute address (what the RE tracer calls
' name_at_wC = rd_bstr(*(node+0xC))). The engine allocates with
' SysAllocStringByteLen, so SysStringLen already yields the character count.
Private Function ReadBstrAt(ByVal p As Long) As String
    Dim n As Long, i As Long, s As String, w As Long
    If p = 0 Then Exit Function
    ' SysStringLen on a non-BSTR pointer faults INSIDE oleaut32; refuse first.
    If IsBadReadPtr(p, 8) <> 0 Then Exit Function
    n = SysStringLen(p)
    If n <= 0 Then Exit Function
    If n > 256 Then n = 256
    For i = 0 To n - 1
        w = 0
        CopyMemory w, p + i * 2, 2
        s = s & ChrW$(w And &HFFFF&)
    Next
    ReadBstrAt = s
End Function

Private Function Fix16(ByVal v As String) As String
    Dim i As Long, out As String
    If InStr(v, Chr$(0)) = 0 Then
        Fix16 = v
        Exit Function
    End If
    out = ""
    For i = 1 To Len(v) Step 2
        out = out & Mid$(v, i, 1)
    Next
    Fix16 = out
End Function

' Immediate-window executor probe. Returns the two engine globals, the raw
' HRESULT and any VB error -- and refuses to call at all when the engine
' singleton is 0 (calling then dereferences null inside VBA6).
' Read-only observation of the two-state
' difference. A correction established that the evaluation frame is the plain global
' dword_FC0D0A4 (RVA 0x17D0A4), cleared while EbMode()==0 and set to the breakpoint's
' frame while broken. This probe only reads: no writes, no EbExecuteLine call.
' Call the engine's OWN frame setter
' sub_FBC3163(a1, a2) at RVA 0x133163 (__stdcall, ret 8). It walks engine+0x2F0
' for the first kind==1 node and then does extra work the raw global write does not
' (sub_FAA3E34(1,1,1), sub_FB9A9FF, ...), so it is the authoritative way to bind the
' evaluation frame. a2=1 means "find the named frame".
Public Function CallFrameSetterJson(ByVal a1 As Long, ByVal a2 As Long) As String
    Dim hMod As Long, target As Long, mem As Long, rc As Long
    Dim b(47) As Byte
    Dim s As String
    hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    If hMod = 0 Then
        CallFrameSetterJson = FailErr(s & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    target = hMod + &H133163
    s = s & Q("target") & ":" & CStr(target)
    s = s & "," & Q("a1") & ":" & CStr(a1) & "," & Q("a2") & ":" & CStr(a2)
    mem = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Then
        CallFrameSetterJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    ' push a2
    b(0) = &H68
    Put32 b, 1, a2
    ' push a1
    b(5) = &H68
    Put32 b, 6, a1
    ' mov eax, target
    b(10) = &HB8
    Put32 b, 11, target
    ' call eax
    b(15) = &HFF
    b(16) = &HD0
    ' ret 16
    b(17) = &HC2
    b(18) = &H10
    b(19) = &H0
    CopyMemoryToPtr mem, b(0), 20
    s = s & "," & Q("thunk") & ":" & CStr(mem)
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(mem, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("ret") & ":" & CStr(rc)
    CallFrameSetterJson = s & "}"
End Function

' Generic "call an internal function by RVA" via the proven thunk mechanism.
' argc must match the callee's own stack cleanup (ret argc*4); the thunk always
' finishes with "ret 16" to consume the four arguments CallWindowProcA forwards.
Public Function CallRvaJson(ByVal rva As Long, ByVal argc As Long, ByVal a1 As Long, ByVal a2 As Long, Optional ByVal hModOverride As Long = 0) As String
    Dim hMod As Long, target As Long, mem As Long, rc As Long
    Dim b(47) As Byte
    Dim p As Long
    Dim s As String
    hMod = hModOverride
    If hMod = 0 Then hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    If hMod = 0 Then
        CallRvaJson = FailErr(s & Q("error") & ":" & Q("module not loaded"), "not_loaded", "module not loaded")
        Exit Function
    End If
    target = hMod + rva
    s = s & Q("rva") & ":" & Q("0x" & Hex$(rva)) & "," & Q("target") & ":" & CStr(target)
    s = s & "," & Q("argc") & ":" & CStr(argc)
    If argc < 0 Or argc > 2 Then
        CallRvaJson = FailErr(s & "," & Q("error") & ":" & Q("argc must be 0..2"), "arg_out_of_range", "argc must be 0..2")
        Exit Function
    End If
    mem = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Then
        CallRvaJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    p = 0
    ' mov ecx, [engine global]  -- many VBA6 internals take the engine in ECX
    ' (thiscall/fastcall). ECX is call-clobbered, so this is harmless when the
    ' callee ignores it.
    If hModOverride = 0 Then b(p) = &H8B: b(p + 1) = &HD: Put32 b, p + 2, hMod + &H17D0F4: p = p + 6
    If argc = 2 Then
        b(p) = &H68: Put32 b, p + 1, a2: p = p + 5
    End If
    If argc >= 1 Then
        b(p) = &H68: Put32 b, p + 1, a1: p = p + 5
    End If
    b(p) = &HB8: Put32 b, p + 1, target: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr mem, b(0), p
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(mem, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number) & "," & Q("vbaErrorDesc") & ":" & Q(Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("ret") & ":" & CStr(rc)
    s = s & "," & Q("retHex") & ":" & Q("0x" & Hex$(rc))
    CallRvaJson = s & "}"
End Function

' The decisive experiment. Call the
' real worker sub_FB7317A (RVA 0xE317A) the way the IDE's own path does:
'   __thiscall, ecx = engine (*(VBA6+0x17D15C)), 8 stack args, retn 20h
'   p0 = ADDRESS of a writable dword that holds the line-text pointer
'   p1..p7 = 1,1,0,0,1,1,0   (IDE truth table)
'   optional: WORD[VBA6+0x17D076] |= 0x100 around the call (IDE does, EbExecuteLine does not)
' The thunk pushes all 8 dwords itself (the worker pops 32 bytes via retn 20h) and
' ends with "ret 16" to consume the four arguments CallWindowProcA forwards.
Public Function WorkerCallJson(ByVal text As String, ByVal asBstr As Boolean, ByVal p1 As Long, ByVal p2 As Long, ByVal p3 As Long, ByVal p4 As Long, ByVal p5 As Long, ByVal p6 As Long, ByVal p7 As Long, ByVal setBit As Boolean, ByVal p2IsEngine As Boolean) As String
    Dim hMod As Long, engine As Long, worker As Long, mem As Long, rc As Long
    Dim tp As Long
    Dim ansi() As Byte
    Dim b(191) As Byte
    Dim p As Long
    Dim wBefore As Long, wSet As Long
    Dim s As String
    hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    If hMod = 0 Then
        WorkerCallJson = FailErr(s & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    engine = 0
    CopyMemory engine, hMod + &H17D15C, 4
    s = s & Q("engine") & ":" & CStr(engine)
    If engine = 0 Then
        WorkerCallJson = FailSkip(s & "," & Q("called") & ":false," & Q("skipped") & ":" & Q("engine singleton is 0"), "engine_zero", "engine singleton is 0")
        Exit Function
    End If
    worker = hMod + &HE317A
    ' Note: when p1 = 0 (the IDE value) the worker may dereference p2
    ' as an object pointer, and p2 = 1 then raises an access violation. The note's
    ' suggested candidate is the engine object itself.
    If p2IsEngine Then p2 = engine
    tp = 0
    If asBstr Then
        tp = SysAllocString(StrPtr(text))
        If tp = 0 Then tp = StrPtr(text)
    Else
        ' The worker reads the ANSI form as a NUL-terminated C string, and StrConv
        ' does NOT append a terminator - without it the worker reads whatever byte
        ' follows the buffer, which made the same call succeed or fail (40036) at
        ' random. Append the terminator explicitly.
        ansi = StrConv(text & vbNullChar, vbFromUnicode)
        If UBound(ansi) >= 0 Then tp = VarPtr(ansi(0))
    End If
    s = s & "," & Q("worker") & ":" & CStr(worker)
    s = s & "," & Q("textPtr") & ":" & CStr(tp) & "," & Q("asBstr") & ":" & B(asBstr)
    mem = VirtualAlloc(0, 512, &H1000, &H40)
    If mem = 0 Then
        WorkerCallJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    ' the writable dword p0 points at lives at mem+400 (the worker may rewrite it)
    CopyMemoryToPtr mem + 400, tp, 4
    p = 0
    b(p) = &H68: Put32 b, p + 1, p7: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p6: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p5: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p4: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p3: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p2: p = p + 5
    b(p) = &H68: Put32 b, p + 1, p1: p = p + 5
    b(p) = &H68: Put32 b, p + 1, mem + 400: p = p + 5
    b(p) = &HB9: Put32 b, p + 1, engine: p = p + 5
    b(p) = &HB8: Put32 b, p + 1, worker: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr mem, b(0), p
    wBefore = 0
    CopyMemory wBefore, hMod + &H17D076, 2
    s = s & "," & Q("scopeBitBefore") & ":" & CStr(wBefore)
    If setBit Then
        wSet = wBefore Or &H100
        CopyMemoryToPtr hMod + &H17D076, wSet, 2
        s = s & "," & Q("scopeBitSet") & ":" & CStr(wSet)
    End If
    s = s & "," & Q("called") & ":true"
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(mem, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number) & "," & Q("vbaErrorDesc") & ":" & Q(Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
    If setBit Then
        CopyMemoryToPtr hMod + &H17D076, wBefore, 2
        s = s & "," & Q("scopeBitRestored") & ":true"
    End If
    s = s & "," & Q("text") & ":" & Q(text)
    s = s & "," & Q("p1") & ":" & CStr(p1) & "," & Q("p2") & ":" & CStr(p2) & "," & Q("p3") & ":" & CStr(p3)
    s = s & "," & Q("p4") & ":" & CStr(p4) & "," & Q("p5") & ":" & CStr(p5) & "," & Q("p6") & ":" & CStr(p6) & "," & Q("p7") & ":" & CStr(p7)
    s = s & "," & Q("ret") & ":" & CStr(rc) & "," & Q("retHex") & ":" & Q("0x" & Hex$(rc))
    WorkerCallJson = s & "}"
End Function

' READ-ONLY memory dump, used to locate
' VBA6's ExFrame stack top and to check ExFrame field offsets. addr takes precedence
' over rva; length is clamped to 512 bytes. Only addresses that look like real
' user-space pointers are accepted - a bad address would fault the IDE.
' ---------------------------------------------------------------------------
' EbGetStatusOfItems(BYTE buf[72], int len).
' VB6.EXE's command dispatcher sub_510336 (VA 0x00510336) calls this on every
' menu command and then tests buf[itemId] to decide whether that command is
' enabled, before handing the id to EbInvokeItem(id). The 72 bytes are thus the
' engine's OWN capability table -- the authoritative list of what an engine can
' be asked to do. Pure read: the callee only fills our buffer.
'   sub_510336:  EbGetStatusOfItems(v34, 72);  ... if (v34[item]) EbInvokeItem(item)
' Resolved by name; RVA 0x20BC is only the fallback (measured evidence).
' ---------------------------------------------------------------------------
Public Function EngineItemsJson() As String
    Dim hMod As Long, fn As Long, byRva As Boolean
    Dim mem As Long, code As Long, p As Long, rc As Long, i As Long, av As Long
    Dim b(31) As Byte
    Dim buf(71) As Byte
    Dim hexs As String
    Dim s As String
    s = "{"
    hMod = Vba6Base()
    If hMod = 0 Then
        EngineItemsJson = FailErr(s & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    fn = GetProcAddress(hMod, "EbGetStatusOfItems")
    byRva = False
    If fn = 0 Then
        fn = hMod + &H20BC
        byRva = True
    End If
    s = s & Q("proc") & ":" & CStr(fn) & "," & Q("viaRva") & ":" & B(byRva)
    mem = VirtualAlloc(0, 128, &H1000, &H40)
    code = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Or code = 0 Then
        EngineItemsJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    ' push len ; push buf ; mov eax,fn ; call eax ; ret 16   (callee retn 8)
    p = 0
    b(p) = &H68: Put32 b, p + 1, 72: p = p + 5
    b(p) = &H68: Put32 b, p + 1, mem: p = p + 5
    b(p) = &HB8: Put32 b, p + 1, fn: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr code, b(0), p
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(code, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("rc") & ":" & CStr(rc)
    s = s & "," & Q("bitmap_len") & ":72"
    If IsBadReadPtr(mem, 72) = 0 Then
        CopyMemory buf(0), mem, 72
        hexs = ""
        av = 0
        For i = 0 To 71
            hexs = hexs & Right$("0" & Hex$(buf(i)), 2)
            If buf(i) <> 0 Then av = av + 1
        Next
        s = s & "," & Q("hex") & ":" & Q(hexs)                 ' unchanged (back-compat)
        s = s & "," & Q("bitmap_hex") & ":" & Q(hexs)
        s = s & "," & Q("byte_offsets") & ":[" & ByteOffsetsJson() & "]"
        s = s & "," & Q("availableCount") & ":" & CStr(av)     ' now documented as NON-ZERO BYTE count
    Else
        s = s & "," & Q("bitmap_hex") & ":null"
        s = s & "," & Q("byte_offsets") & ":null"
        s = s & "," & Q("availableCount") & ":null"
    End If
    s = s & "," & Q("items") & ":null"
    s = s & "," & Q("available") & ":null"
    s = s & "," & Q("confidence") & ":" & Q("inferred")
    s = s & "," & Q("note") & ":" & Q("72-byte status bitmap; no item index semantics")
    s = s & "," & Q("observedTwoStateChanged") & ":[9,20,21,29]"
    s = s & "," & Q("twoStateProbe") & ":" & Q("design vs RUN in one IDE session: +0x09/+0x14/+0x15/+0x1D changed; +0x07/+0x08 stayed 0, so the EbMode==1 gating claim was NOT confirmed")
    EngineItemsJson = s & "}"
End Function

Private Function ByteOffsetsJson() As String
    ' offsets are INFERRED from static write-site analysis (docs/RE_ROUND451_ANSWERS.md G4 r3-r5)
    ByteOffsetsJson = _
        "{""offset"":7,""offsetHex"":""0x07"",""note"":""INFERRED_STATIC +0x07/+0x08: static analysis suggested EbMode==1 gating; the two-state probe did NOT confirm it (both 0 in design and RUN)""}," & _
        "{""offset"":16,""offsetHex"":""0x10"",""note"":""INFERRED_STATIC +0x10/+0x11 derived from the state dword at [0xFC0D0F4]""}," & _
        "{""offset"":19,""offsetHex"":""0x13"",""note"":""INFERRED_STATIC +0x13 state dword low byte (edx-derived)""}," & _
        "{""offset"":30,""offsetHex"":""0x1E"",""note"":""INFERRED_STATIC +0x1E/+0x1F/+0x21 helper/branch-derived status bits""}," & _
        "{""offset"":43,""offsetHex"":""0x2B"",""note"":""INFERRED_STATIC +0x2B normalized result of 0xFA922BE""}"
End Function

' ---------------------------------------------------------------------------
' EbGetCallstackFunction, called exactly the way VB6.EXE's
' OWN caller does it (sub_55D10D, VA 0x0055D10D). An explicit thunk is used
' instead of a VB Declare so there is no marshalling ambiguity: the three BSTR*
' out-slots and arg5 live in memory WE allocated and zeroed, byte-for-byte the
' IDE's protocol --
'     BSTR outA = 0, outB = 0, outC = 0;  int arg5;
'     rc = EbGetCallstackFunction(frame, &outA, &outB, &outC, &arg5);
'     if (!rc) { ... use the strings ... }        <- 0 means SUCCESS
' The CALLEE allocates the BSTRs with SysAllocString, so the caller must free
' them; the IDE does SysFreeString on each.
' CRITICAL: only ever call with 0 <= frame <= count-1. The R2 crash was an
' OUT-OF-RANGE index (mov eax,[edi+4] with edi=NULL), not a broken function --
' A correction narrowed the valid range from 0..count-2 to 0..count-1.
' No VB Declare, no ByRef String: nothing VB can silently marshal.
' ---------------------------------------------------------------------------
Public Function CallstackFunctionJson(ByVal frameIndex As Long) As String
    Dim hMod As Long, fn As Long, byRva As Boolean
    Dim mem As Long, code As Long, p As Long, i As Long, rc As Long, z As Long
    Dim b(47) As Byte
    Dim pA As Long, pB As Long, pC As Long, pArg As Long
    Dim vA As Long, vB As Long, vC As Long, vArg As Long
    Dim sA As String, sB As String, sC As String, joined As String
    Dim s As String
    s = "{"
    hMod = Vba6Base()
    If hMod = 0 Then
        CallstackFunctionJson = FailErr(s & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    fn = GetProcAddress(hMod, "EbGetCallstackFunction")
    byRva = False
    If fn = 0 Then
        fn = hMod + &H13321A
        byRva = True
    End If
    s = s & Q("frame") & ":" & CStr(frameIndex)
    s = s & "," & Q("proc") & ":" & CStr(fn) & "," & Q("viaRva") & ":" & B(byRva)
    ' R23 搂2.1: the valid range is 0..count-1. Refuse anything else -- an
    ' out-of-range index is exactly what crashed the IDE before.
    Dim cnt As Long
    cnt = -1
    Call EbGetCallstackCount(cnt)
    s = s & "," & Q("callstackCount") & ":" & CStr(cnt)
    If cnt <= 0 Or frameIndex < 0 Or frameIndex > cnt - 1 Then
        CallstackFunctionJson = FailSkip(s & "," & Q("skipped") & ":" & Q("frame index out of range; the engine was NOT called"), "arg_out_of_range", "frame index out of range; the engine was NOT called")
        Exit Function
    End If
    mem = VirtualAlloc(0, 64, &H1000, &H40)
    code = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Or code = 0 Then
        CallstackFunctionJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    ' the IDE's first line: all three out-slots AND arg5 start at zero
    z = 0
    For i = 0 To 7
        CopyMemoryToPtr mem + i * 4, z, 4
    Next
    pA = mem
    pB = mem + 4
    pC = mem + 8
    pArg = mem + &H20
    ' __stdcall pushes RIGHT to LEFT: arg5, outC, outB, outA, frame
    p = 0
    b(p) = &H68: Put32 b, p + 1, pArg: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pC: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pB: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pA: p = p + 5
    b(p) = &H68: Put32 b, p + 1, frameIndex: p = p + 5
    b(p) = &HB8: Put32 b, p + 1, fn: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr code, b(0), p
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(code, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("rc") & ":" & CStr(rc) & "," & Q("retHex") & ":" & Q("0x" & Hex$(rc))
    ' On failure the out-slots are not guaranteed, so only read them on rc = 0.
    If rc = 0 Then
        CopyMemory vA, pA, 4
        CopyMemory vB, pB, 4
        CopyMemory vC, pC, 4
        CopyMemory vArg, pArg, 4
        s = s & "," & Q("outA") & ":" & CStr(vA) & "," & Q("outB") & ":" & CStr(vB)
        s = s & "," & Q("outC") & ":" & CStr(vC) & "," & Q("arg5Out") & ":" & CStr(vArg)
        sA = ReadBstrAt(vA)
        sB = ReadBstrAt(vB)
        sC = ReadBstrAt(vC)
        s = s & "," & Q("strA") & ":" & Q(sA)
        s = s & "," & Q("strB") & ":" & Q(sB)
        s = s & "," & Q("strC") & ":" & Q(sC)
        ' the IDE's own format: segment, then '.' + segment for each non-empty one
        joined = sA
        If Len(sB) > 0 Then joined = joined & "." & sB
        If Len(sC) > 0 Then joined = joined & "." & sC
        s = s & "," & Q("joined") & ":" & Q(joined)
        ' the callee allocated these -- the caller owns them
        If vA <> 0 Then SysFreeString vA
        If vB <> 0 Then SysFreeString vB
        If vC <> 0 Then SysFreeString vC
    End If
    CallstackFunctionJson = s & "}"
End Function
' Shared low-level frame read used by the callstack TOOL (CallstackJson).
' CallstackFunctionJson above deliberately keeps its OWN copy of the thunk: it
' is the single-frame RE probe that was verified in a live break session, so it
' stays byte-identical to what was verified.
Private Function CsCallOne(ByVal frameIndex As Long, ByRef vA As Long, ByRef vB As Long, ByRef vC As Long, ByRef vArg As Long, ByRef viaRva As Boolean) As Long
    Dim hMod As Long, fn As Long
    Dim mem As Long, code As Long, p As Long, i As Long, z As Long
    Dim b(47) As Byte
    Dim pA As Long, pB As Long, pC As Long, pArg As Long
    hMod = Vba6Base()
    If hMod = 0 Then Exit Function
    fn = GetProcAddress(hMod, "EbGetCallstackFunction")
    viaRva = False
    If fn = 0 Then
        fn = hMod + &H13321A
        viaRva = True
    End If
    mem = VirtualAlloc(0, 64, &H1000, &H40)
    code = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Or code = 0 Then Exit Function
    ' the IDE's protocol: all three BSTR* out-slots AND arg5 start at zero
    z = 0
    For i = 0 To 7
        CopyMemoryToPtr mem + i * 4, z, 4
    Next
    pA = mem
    pB = mem + 4
    pC = mem + 8
    pArg = mem + &H20
    ' __stdcall pushes RIGHT to LEFT: arg5, outC, outB, outA, frame
    p = 0
    b(p) = &H68: Put32 b, p + 1, pArg: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pC: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pB: p = p + 5
    b(p) = &H68: Put32 b, p + 1, pA: p = p + 5
    b(p) = &H68: Put32 b, p + 1, frameIndex: p = p + 5
    b(p) = &HB8: Put32 b, p + 1, fn: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr code, b(0), p
    z = 0
    On Error Resume Next
    z = CallWindowProcA(code, 0, 0, 0, 0)
    Err.Clear
    On Error GoTo 0
    ' rc = 0 is SUCCESS; on failure the out-slots are not guaranteed
    If z = 0 Then
        CopyMemory vA, pA, 4
        CopyMemory vB, pB, 4
        CopyMemory vC, pC, 4
        CopyMemory vArg, pArg, 4
    End If
    CsCallOne = z
End Function

' "project.module.proc" -- the IDE's own format (sub_55D10D joins the three
' returned BSTRs with '.', skipping empty ones).
Private Function CsJoin(ByVal a As String, ByVal b As String, ByVal c As String) As String
    Dim j As String
    j = a
    If Len(b) > 0 Then
        If Len(j) > 0 Then j = j & "." & b Else j = b
    End If
    If Len(c) > 0 Then
        If Len(j) > 0 Then j = j & "." & c Else j = c
    End If
    CsJoin = j
End Function

' ---------------------------------------------------------------------------
' G1 -- per-frame PROCEDURE NAME, as a shippable tool (tool: get_callstack).
' Every frame of the IDE's own call-stack window is reproduced here: the engine
' returns the project name, the MODULE and the PROCEDURE for each frame.
' Verified live: at a breakpoint in modMain.Probe
' called from modMain.Main the result was
'     [0] brk.modMain.Probe      [1] brk.modMain.Main       count = 2
' SAFETY: EbGetCallstackCount/Function are only valid while the program is
' stopped. This refuses to touch the engine unless EbMode() = 2 (break), and
' never indexes past count-1 -- an out-of-range index is what crashed the IDE
' earlier, not a defect in the function itself.
' ---------------------------------------------------------------------------
Public Function CallstackJson(ByVal maxFrames As Long) As String
    Dim mode As Long, cnt As Long, i As Long, n As Long, rc As Long
    Dim vA As Long, vB As Long, vC As Long, vArg As Long
    Dim viaRva As Boolean
    Dim fA As String, fB As String, fC As String
    Dim frames As String, s As String
    s = "{"
    mode = EbMode()
    s = s & Q("ebMode") & ":" & CStr(mode)
    If mode <> 2 Then
        CallstackJson = s & "," & Q("broken") & ":false," & Q("count") & ":0," & Q("frameCount") & ":0" _
            & "," & Q("frames") & ":[]" _
            & "," & Q("note") & ":" & Q("the program is not stopped at a breakpoint, so there is no call stack") & "}"
        Exit Function
    End If
    cnt = -1
    Call EbGetCallstackCount(cnt)
    s = s & "," & Q("broken") & ":true," & Q("count") & ":" & CStr(cnt)
    If maxFrames <= 0 Then maxFrames = 32
    If maxFrames > 128 Then maxFrames = 128
    n = cnt
    If n > maxFrames Then n = maxFrames
    If n < 0 Then n = 0
    frames = ""
    For i = 0 To n - 1
        vA = 0
        vB = 0
        vC = 0
        vArg = 0
        rc = CsCallOne(i, vA, vB, vC, vArg, viaRva)
        fA = ""
        fB = ""
        fC = ""
        If rc = 0 Then
            fA = ReadBstrAt(vA)
            fB = ReadBstrAt(vB)
            fC = ReadBstrAt(vC)
        End If
        If Len(frames) > 0 Then frames = frames & ","
        frames = frames & "{" & Q("index") & ":" & CStr(i) & "," & Q("rc") & ":" & CStr(rc) _
            & "," & Q("project") & ":" & Q(fA) & "," & Q("module") & ":" & Q(fB) _
            & "," & Q("proc") & ":" & Q(fC) & "," & Q("full") & ":" & Q(CsJoin(fA, fB, fC)) & "}"
        If vA <> 0 Then SysFreeString vA
        If vB <> 0 Then SysFreeString vB
        If vC <> 0 Then SysFreeString vC
    Next
    s = s & "," & Q("frameCount") & ":" & CStr(n)
    s = s & "," & Q("frames") & ":[" & frames & "]"
    s = s & "," & Q("viaRva") & ":" & B(viaRva)
    CallstackJson = s & "}"
End Function
' ---------------------------------------------------------------------------
' Verification aid: report the REAL RVA of the VBA6 exports the
' RE names, resolved BY NAME through the real PE export table. This turns every
' "function X lives at RVA 0xNNNN" claim into a single safe (read-only) call.
' ---------------------------------------------------------------------------
Public Function ExportRvasJson() As String
    Dim names As Variant, i As Long, p As Long, hMod As Long, rva As Long
    Dim s As String
    names = Array("EbGetStatusOfItems", "EbInvokeItem", "EbGetCallstackFunction", _
                  "EbGetCallstackCount", "EbGetDebWndCallstackIndex", "EbIsWatchExprOk", _
                  "EbWatchDlgModSelected", "EbIsValidIdent", "EbGetCurrentFunction", _
                  "EbMode", "EbExecuteLine", "EbGetVBAObject", "EbNotifyWindow", _
                  "EbSetStepToCursor", "EbSetStepMode", "EbGetStepMode", "EbGetOptions", _
                  "EbIsCodeOnStack", "EbShowFunction", "EbExtShowEventProc")
    hMod = Vba6Base()
    s = "{"
    s = s & Q("vba6Base") & ":" & CStr(hMod)
    If hMod = 0 Then
        ExportRvasJson = FailErr(s & "," & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    s = s & "," & Q("exports") & ":["
    For i = 0 To UBound(names)
        p = GetProcAddress(hMod, CStr(names(i)))
        rva = 0
        If p <> 0 Then rva = p - hMod
        If i > 0 Then s = s & ","
        s = s & "{" & Q("name") & ":" & Q(CStr(names(i))) & "," & Q("found") & ":" & B(p <> 0) _
            & "," & Q("addr") & ":" & CStr(p) & "," & Q("rva") & ":" & Q("0x" & Hex$(rva)) & "}"
    Next
    s = s & "]"
    ExportRvasJson = s & "}"
End Function
' ---------------------------------------------------------------------------
' Verification aid: a READ-ONLY dword scanner over a live range.
' R32 headline is a NEGATIVE claim -- VB6.EXE contains no absolute constant that
' points into the VBA6 engine global area 0x0FC0D000..0x0FC0F000. Replicating a
' whole-image scan from the wire would need thousands of mem_dump calls, so the
' scan runs INSIDE the process. It is BYTE-WISE, not just 4-aligned, because
' R32 decisive argument is that its two hits were MISALIGNED instruction bytes
' rather than real references -- an aligned-only scan could not test that.
'   vba6_probe { mem_scan: true, scan_addr, scan_len, scan_lo, scan_hi, scan_max }
' ---------------------------------------------------------------------------
Public Function MemScanJson(ByVal addr As Long, ByVal scanLen As Long, ByVal lo As Long, ByVal hi As Long, ByVal maxHits As Long) As String
    Dim i As Long, n As Long, off As Long
    Dim v As Double, loD As Double, hiD As Double
    ' NB: this array must NOT be named 'b' -- VB6 is case-insensitive and 'b'
    ' would collide with the helper B() used below for the alignment flag
    ' (measured: it silently indexed the buffer and printed the first byte).
    Dim buf() As Byte
    Dim hits As String, s As String
    Dim lenReq As Long, lenClamped As Boolean, lenDefaulted As Boolean
    Dim hitsReq As Long, hitsClamped As Boolean, hitsDefaulted As Boolean
    s = "{"
    lenReq = scanLen
    hitsReq = maxHits
    If scanLen <= 0 Then scanLen = &H10000
    lenDefaulted = (lenReq <= 0)
    If scanLen > &H100000 Then scanLen = &H100000
    lenClamped = (lenReq > &H100000)
    If maxHits <= 0 Then maxHits = 64
    hitsDefaulted = (hitsReq <= 0)
    If maxHits > 512 Then maxHits = 512
    hitsClamped = (hitsReq > 512)
    s = s & Q("addr") & ":" & CStr(addr) & "," & Q("len") & ":" & CStr(scanLen)
    s = s & "," & Q("lenRequested") & ":" & CStr(lenReq) & "," & Q("lenDefaulted") & ":" & B(lenDefaulted) & "," & Q("lenClamped") & ":" & B(lenClamped)
    s = s & "," & Q("maxHits") & ":" & CStr(maxHits) & "," & Q("maxHitsRequested") & ":" & CStr(hitsReq) & "," & Q("maxHitsDefaulted") & ":" & B(hitsDefaulted) & "," & Q("maxHitsClamped") & ":" & B(hitsClamped)
    s = s & "," & Q("lo") & ":" & CStr(lo) & "," & Q("hi") & ":" & CStr(hi)
    If addr < &H10000 Or addr > &H7FFFFFFF Then
        MemScanJson = FailErr(s & "," & Q("error") & ":" & Q("address looks invalid; refusing to read"), "address_invalid", "address looks invalid; refusing to read")
        Exit Function
    End If
    If lo = 0 Then
        MemScanJson = FailErr(s & "," & Q("error") & ":" & Q("lo must be non-zero, every zero dword would match"), "arg_out_of_range", "lo must be non-zero, every zero dword would match")
        Exit Function
    End If
    If IsBadReadPtr(addr, scanLen) <> 0 Then
        MemScanJson = FailErr(s & "," & Q("error") & ":" & Q("range is not fully readable"), "not_readable", "range is not fully readable")
        Exit Function
    End If
    ReDim buf(scanLen - 1)
    CopyMemory buf(0), addr, scanLen
    loD = CDbl(lo)
    hiD = CDbl(hi)
    n = 0
    hits = ""
    For i = 0 To scanLen - 4
        v = CDbl(buf(i)) + CDbl(buf(i + 1)) * 256# + CDbl(buf(i + 2)) * 65536# + CDbl(buf(i + 3)) * 16777216#
        If v >= loD And v <= hiD Then
            n = n + 1
            If n <= maxHits Then
                off = addr + i
                If Len(hits) > 0 Then hits = hits & ","
                hits = hits & "{" & Q("at") & ":" & CStr(off) & "," & Q("atHex") & ":" & Q("0x" & Hex$(off)) _
                    & "," & Q("aligned") & ":" & B((off And 3) = 0) & "," & Q("v") & ":" & Q("0x" & Hex$(CLng(v))) & "}"
            End If
        End If
    Next
    s = s & "," & Q("hitCount") & ":" & CStr(n)
    s = s & "," & Q("hits") & ":[" & hits & "]"
    MemScanJson = s & "}"
End Function

' ---------------------------------------------------------------------------
' Follow-up: WHO OWNS THE IDE WINDOWS?
' R32/R33 prove VB6.EXE never touches VBA6 memory and holds no engine COM
' pointer, and conclude the debug windows must be fed through an export or a
' callback. That reasoning has one unexamined alternative: the debug windows
' might be created and populated BY VBA6.DLL ITSELF, in which case VB6.EXE needs
' neither engine memory nor a locals export.
' Cross-process GetWindowLong(GWL_WNDPROC) and GetClassLong(GCL_HMODULE) both
' return 0, so the ownership test MUST run in-process (that is this function).
' For every window of this process it reports the window procedure AND the
' module that REGISTERED the window class, each classified by module range.
' ---------------------------------------------------------------------------

Private Function WndText(ByVal h As Long, ByVal asClass As Boolean) As String
    Dim buf As String, n As Long, z As Long
    On Error Resume Next
    buf = String$(300, Chr$(0))
    If asClass Then
        n = GetClassNameW(h, StrPtr(buf), 300)
    Else
        n = GetWindowTextW(h, StrPtr(buf), 300)
    End If
    If n <= 0 Then Exit Function
    z = InStr(buf, Chr$(0))
    If z > 0 Then buf = Left$(buf, z - 1)
    WndText = buf
    Err.Clear
End Function

' Classify an address by which image it falls in (both images are non-relocating).
Private Function AddrModule(ByVal addr As Long) As String
    Dim hb As Long, vb As Long, hs As Long, vs As Long
    If addr = 0 Then
        AddrModule = "null"
        Exit Function
    End If
    hb = HostBase()
    vb = Vba6Base()
    hs = PeImageSize(hb)
    vs = PeImageSize(vb)
    If hs > 0 And addr >= hb And addr < hb + hs Then
        AddrModule = "vb6"
    ElseIf vs > 0 And addr >= vb And addr < vb + vs Then
        AddrModule = "vba6"
    Else
        AddrModule = "other"
    End If
End Function

Public Function WindowProbeJson(ByVal maxN As Long) As String
    Dim h As Long, c As Long, n As Long, i As Long
    Dim wp As Long, cm As Long
    Dim cls As String, ttl As String, wm As String, cmm As String
    Dim items As String, s As String
    Dim pid As Long, wpid As Long
    Dim histVb6 As Long, histVba6 As Long, histOther As Long, histNull As Long
    Dim classMods As String
    If maxN <= 0 Then maxN = 200
    If maxN > 1500 Then maxN = 1500
    pid = GetCurrentProcessId()
    s = "{"
    s = s & Q("pid") & ":" & CStr(pid)
    s = s & "," & Q("hostBase") & ":" & CStr(HostBase())
    s = s & "," & Q("vba6Base") & ":" & CStr(Vba6Base())
    items = ""
    classMods = ""
    n = 0
    histVb6 = 0
    histVba6 = 0
    histOther = 0
    histNull = 0
    ' walk the whole top-level list, keep only this process, then descend one level
    h = GetWindow(GetDesktopWindow(), 5)
    i = 0
    Do While h <> 0 And i < 4000
        i = i + 1
        wpid = 0
        Call GetWindowThreadProcessId(h, wpid)
        If wpid = pid Then
            c = GetWindow(h, 5)
            Do While c <> 0 And n < maxN
                n = n + 1
                cls = WndText(c, True)
                ttl = WndText(c, False)
                wp = GetWindowLongW(c, -4)
                cm = GetClassLongW(c, -16)
                wm = AddrModule(wp)
                cmm = AddrModule(cm)
                If cmm = "vb6" Then
                    histVb6 = histVb6 + 1
                ElseIf cmm = "vba6" Then
                    histVba6 = histVba6 + 1
                ElseIf cmm = "other" Then
                    histOther = histOther + 1
                Else
                    histNull = histNull + 1
                End If
                If InStr(classMods, "|" & cls & "=" & cmm & "|") = 0 Then
                    classMods = classMods & "|" & cls & "=" & cmm & "|"
                End If
                If Len(items) > 0 Then items = items & ","
                items = items & "{" & Q("hwnd") & ":" & CStr(c) & "," & Q("cls") & ":" & Q(cls) _
                    & "," & Q("title") & ":" & Q(ttl) & "," & Q("wndProc") & ":" & CStr(wp) _
                    & "," & Q("wndProcMod") & ":" & Q(wm) & "," & Q("classMod") & ":" & Q(cmm) & "}"
                c = GetWindow(c, 2)
            Loop
        End If
        h = GetWindow(h, 2)
    Loop
    s = s & "," & Q("count") & ":" & CStr(n)
    s = s & "," & Q("byClassMod") & ":{" & Q("vb6") & ":" & CStr(histVb6) _
        & "," & Q("vba6") & ":" & CStr(histVba6) & "," & Q("other") & ":" & CStr(histOther) _
        & "," & Q("none") & ":" & CStr(histNull) & "}"
    s = s & "," & Q("windows") & ":[" & items & "]"
    s = s & "," & Q("classModules") & ":" & Q(classMods)
    WindowProbeJson = s & "}"
End Function

' ---------------------------------------------------------------------------
' NEW DIRECTION (MCP side): the engine's HOSTED expression service.
' CreateHostedExprSrv (RVA 0x14A3E9) is one of the 317 VBA6 exports that VB6.EXE
' NEVER imports. Every RE round asked "what does the IDE call?", so this function
' is structurally invisible to that method -- yet MCP, running inside the process,
' can call it directly. Evidence (read-only, verified live):
'   CreateHostedExprSrv(void **ppv)  -- 1 arg, out pointer, zeroes *ppv, then
'   allocates 404 bytes (push 194h) and constructs the object;
'   the constructor installs vtable 0x0FB437A0, whose slot run is exactly 7
'   (terminator dword[7] = 0).
' The same technique reproduced R22 known IExprSrv vtable 0x0FAF1F50 from the
' body of CreateIExprSrvObj, which is what validates it.
' This probe instantiates the service so the object and its 7 slots can be read.
' ---------------------------------------------------------------------------
Public Function HostedExprSrvJson() As String
    Dim hMod As Long, fn As Long, viaRva As Boolean
    Dim mem As Long, code As Long, p As Long, rc As Long, i As Long, z As Long
    Dim b(47) As Byte
    Dim pOut As Long, pEsp As Long, obj As Long, vt As Long, sv As Long
    Dim slots As String, s As String
    s = "{"
    hMod = Vba6Base()
    If hMod = 0 Then
        HostedExprSrvJson = FailErr(s & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    fn = GetProcAddress(hMod, "CreateHostedExprSrv")
    viaRva = False
    If fn = 0 Then
        fn = hMod + &H14A3E9
        viaRva = True
    End If
    s = s & Q("proc") & ":" & CStr(fn) & "," & Q("viaRva") & ":" & B(viaRva)
    mem = VirtualAlloc(0, 64, &H1000, &H40)
    code = VirtualAlloc(0, 64, &H1000, &H40)
    If mem = 0 Or code = 0 Then
        HostedExprSrvJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    ' out slot at +0, saved-esp slot at +4
    z = 0
    CopyMemoryToPtr mem, z, 4
    CopyMemoryToPtr mem + 4, z, 4
    pOut = mem
    pEsp = mem + 4
    ' ARITY-INDEPENDENT thunk -- this is the part that makes the call safe:
    '   mov [pEsp], esp ; push pOut ; mov eax, fn ; call eax
    '   mov esp, [pEsp] ; ret 16
    ' Restoring esp ourselves makes the callee OWN stack cleanup irrelevant, so an
    ' unknown retn N (measured: the first retn in the body is 8, not 4) can no
    ' longer unbalance the caller. It also removes any cdecl/stdcall assumption.
    p = 0
    b(p) = &H89: b(p + 1) = &H25: Put32 b, p + 2, pEsp: p = p + 6
    b(p) = &H68: Put32 b, p + 1, pOut: p = p + 5
    b(p) = &HB8: Put32 b, p + 1, fn: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &H8B: b(p + 1) = &H25: Put32 b, p + 2, pEsp: p = p + 6
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    s = s & "," & Q("espRestore") & ":true"
    CopyMemoryToPtr code, b(0), p
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(code, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("rc") & ":" & CStr(rc)
    obj = 0
    If IsBadReadPtr(pOut, 4) = 0 Then CopyMemory obj, pOut, 4
    s = s & "," & Q("obj") & ":" & CStr(obj)
    If obj > &H10000 Then
        If IsBadReadPtr(obj, 4) = 0 Then
            vt = 0
            CopyMemory vt, obj, 4
            s = s & "," & Q("vtable") & ":" & CStr(vt)
            slots = ""
            For i = 0 To 9
                sv = 0
                If IsBadReadPtr(vt + i * 4, 4) = 0 Then CopyMemory sv, vt + i * 4, 4
                If Len(slots) > 0 Then slots = slots & ","
                slots = slots & CStr(sv)
            Next
            s = s & "," & Q("slots") & ":[" & slots & "]"
        End If
    End If
    HostedExprSrvJson = s & "}"
End Function
Public Function MemDumpJson(ByVal rva As Long, ByVal addr As Long, ByVal length As Long, Optional ByVal hModOverride As Long = 0) As String
    Dim hMod As Long, target As Long, i As Long, v As Long
    Dim b() As Byte
    Dim s As String, hexs As String, dwo As String
    hMod = hModOverride
    If hMod = 0 Then hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    If hMod = 0 Then
        MemDumpJson = FailErr(s & Q("error") & ":" & Q("module not loaded"), "not_loaded", "module not loaded")
        Exit Function
    End If
    If addr <> 0 Then
        target = addr
    Else
        target = hMod + rva
    End If
    s = s & Q("modBase") & ":" & CStr(hMod)
    s = s & "," & Q("rva") & ":" & Q("0x" & Hex$(rva))
    s = s & "," & Q("target") & ":" & CStr(target)
    If target < &H10000 Or target > &H7FFFFFFF Then
        MemDumpJson = FailErr(s & "," & Q("error") & ":" & Q("address looks invalid; refusing to read"), "address_invalid", "address looks invalid; refusing to read")
        Exit Function
    End If
    ' measured: 0x00100000 passes the numeric check and still kills the IDE
    If IsBadReadPtr(target, 4) <> 0 Then
        MemDumpJson = FailErr(s & "," & Q("error") & ":" & Q("address is not readable; refusing to read"), "not_readable", "address is not readable; refusing to read")
        Exit Function
    End If
    Dim lenReq As Long, lenClamped As Boolean, lenDefaulted As Boolean
    lenReq = length
    If length <= 0 Then length = 64
    lenDefaulted = (lenReq <= 0)
    If length > 512 Then length = 512
    lenClamped = (lenReq > 512)
    length = (length \ 4) * 4
    If length = 0 Then length = 4
    ReDim b(length - 1)
    CopyMemory b(0), target, length
    hexs = ""
    dwo = ""
    For i = 0 To length - 1
        hexs = hexs & Right$("0" & Hex$(b(i)), 2)
    Next
    For i = 0 To (length \ 4) - 1
        v = 0
        CopyMemory v, target + i * 4, 4
        If i > 0 Then dwo = dwo & ","
        dwo = dwo & "{" & Q("off") & ":" & Q("0x" & Hex$(i * 4)) & "," & Q("d") & ":" & CStr(v) & "," & Q("hex") & ":" & Q("0x" & Hex$(v)) & "}"
    Next
    s = s & "," & Q("len") & ":" & CStr(length)
    s = s & "," & Q("lenRequested") & ":" & CStr(lenReq) & "," & Q("lenDefaulted") & ":" & B(lenDefaulted) & "," & Q("lenClamped") & ":" & B(lenClamped)
    s = s & "," & Q("hex") & ":" & Q(hexs)
    s = s & "," & Q("dwords") & ":[" & dwo & "]"
    MemDumpJson = s & "}"
End Function

Private Sub Put32(ByRef b() As Byte, ByVal off As Long, ByVal v As Long)
    b(off) = v And &HFF
    b(off + 1) = (v \ &H100) And &HFF
    b(off + 2) = (v \ &H10000) And &HFF
    b(off + 3) = (v \ &H1000000) And &HFF
End Sub

Public Function FrameStateJson() As String
    Dim hMod As Long, engine As Long, frameGlobal As Long
    Dim pHead As Long, node As Long, nNext As Long, nKind As Long, payload As Long
    Dim k As Long, nodes As String
    Dim s As String
    hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    s = s & Q("ebMode") & ":" & CStr(EbMode())
    s = s & "," & ModuleInfoJson()
    If hMod = 0 Then
        FrameStateJson = FailErr(s & "," & Q("error") & ":" & Q("VBA6.DLL not loaded"), "not_loaded", "VBA6.DLL not loaded")
        Exit Function
    End If
    engine = 0
    CopyMemory engine, hMod + &H17D0F4, 4
    frameGlobal = 0
    CopyMemory frameGlobal, hMod + &H17D0A4, 4
    s = s & "," & Q("engineAt17D0F4") & ":" & CStr(engine)
    s = s & "," & Q("evalFrameGlobal") & ":" & CStr(frameGlobal)
    s = s & "," & Q("evalFrameIsZero") & ":" & B(frameGlobal = 0)
    If engine = 0 Then
        FrameStateJson = s & "}"
        Exit Function
    End If
    pHead = 0
    CopyMemory pHead, engine + &H2F0, 4
    s = s & "," & Q("frameChainHead") & ":" & CStr(pHead)
    node = pHead
    nodes = ""
    For k = 0 To 15
        If node = 0 Then Exit For
        CopyMemory nNext, node, 4
        CopyMemory nKind, node + 4, 4
        CopyMemory payload, node + 8, 4
        If Len(nodes) > 0 Then nodes = nodes & ","
        nodes = nodes & "{" & Q("i") & ":" & CStr(k) & "," & Q("node") & ":" & CStr(node) _
            & "," & Q("kind") & ":" & CStr(nKind) & "," & Q("payload") & ":" & CStr(payload) & "}"
        node = nNext
    Next
    s = s & "," & Q("chain") & ":[" & nodes & "]"

    ' -------------------------------------------------------------------------
    ' An earlier batch: locate the current function's ebp.
    ' Locals/args are addressed as SIGNED 16-bit offsets from ebp (negative =
    ' locals, positive = args); the function's RTMI lives at [ebp-0x50] and
    ' housekeeping occupies ebp-0x94..ebp. ebp is not stored in any ExFrame field
    ' we have identified, so scan a window around the current frame for
    ' candidates whose [e-0x50] has the RTMI shape: readable, and *RTMI readable,
    ' and *(*RTMI+4) (ObjectInfo.lpObjectTable) readable.
    ' The cross-check is [ExFrame+0x0C], which VBAStack reads as RTMI.
    ' -------------------------------------------------------------------------
    Dim curFrame As Long, exFrame As Long, rtmiFromEx As Long
    curFrame = 0
    CopyMemory curFrame, hMod + &H17D00C, 4
    exFrame = 0
    If curFrame <> 0 Then
        If IsBadReadPtr(curFrame, 4) = 0 Then CopyMemory exFrame, curFrame, 4
    End If
    rtmiFromEx = 0
    If exFrame <> 0 Then
        If IsBadReadPtr(exFrame + &HC, 4) = 0 Then CopyMemory rtmiFromEx, exFrame + &HC, 4
    End If
    s = s & "," & Q("curFrameAt17D00C") & ":" & CStr(curFrame)
    s = s & "," & Q("exFrame") & ":" & CStr(exFrame)
    s = s & "," & Q("rtmiFromExFrame0C") & ":" & CStr(rtmiFromEx)
    ' The frame global may point AT the ExFrame rather than at a wrapper (measured:
    ' *(curFrame) reads 0 in a clean run). Dump the frame's own dwords so the layout
    ' can be read off directly instead of assumed.
    Dim fd0 As Long, fd4 As Long, fd8 As Long, fdC As Long
    Dim fd10 As Long, fd14 As Long, fd18 As Long, fd1C As Long
    fd0 = 0: fd4 = 0: fd8 = 0: fdC = 0
    fd10 = 0: fd14 = 0: fd18 = 0: fd1C = 0
    If curFrame > &H10000 Then
        If IsBadReadPtr(curFrame, &H20) = 0 Then
            CopyMemory fd0, curFrame, 4
            CopyMemory fd4, curFrame + 4, 4
            CopyMemory fd8, curFrame + 8, 4
            CopyMemory fdC, curFrame + &HC, 4
            CopyMemory fd10, curFrame + &H10, 4
            CopyMemory fd14, curFrame + &H14, 4
            CopyMemory fd18, curFrame + &H18, 4
            CopyMemory fd1C, curFrame + &H1C, 4
        End If
    End If
    s = s & "," & Q("frameDwords") & ":[" & CStr(fd0) & "," & CStr(fd4) & "," & CStr(fd8) & "," & CStr(fdC) & "," & CStr(fd10) & "," & CStr(fd14) & "," & CStr(fd18) & "," & CStr(fd1C) & "]"
    ' Reference RTMI for the cross-check: prefer [ExFrame+0x0C], else [curFrame+0x0C].
    Dim rtmiRef As Long
    If exFrame <> 0 Then
        rtmiRef = rtmiFromEx
    Else
        rtmiRef = fdC
    End If
    s = s & "," & Q("rtmiRef") & ":" & CStr(rtmiRef)

    ' ---- value-driven localisation (an earlier batch) -------------------------------------
    ' The RTMI-shape heuristic found no match, so do not guess ebp: at this
    ' breakpoint the Form_Load locals are KNOWN (n = 42 as Long, s = "hello").
    ' A local slot is a 16-byte VARIANT: [+0] = type code, [+8] = data.
    ' Search the window for those exact values instead.
    Dim scanLo As Long, p As Long
    Dim vt As Long, dp As Long, w0 As Long, w1 As Long
    Dim hits As String, nh As Long, tag As String
    hits = ""
    nh = 0
    If curFrame > &H2000 Then
        scanLo = curFrame - &H8000
        For p = scanLo To curFrame + &H8000 Step 4
            If p > &H10000 Then
                If IsBadReadPtr(p, &H10) = 0 Then
                    vt = 0
                    CopyMemory vt, p, 2
                    dp = 0
                    CopyMemory dp, p + 8, 4
                    tag = ""
                    ' (a) variant slot holding Long 42, and (b) the raw dword 42
                    ' anywhere (native layout) are both reported.
                    If vt = 3 And dp = 42 Then tag = "slot:n42"
                    If dp = 42 Then tag = tag & "+raw42"
                    ' (c) any dword that points at a full UTF-16 "hello" BSTR
                    If dp > &H10000 And dp < &H7FFFFFFF Then
                        If IsBadReadPtr(dp, 12) = 0 Then
                            w0 = 0
                            CopyMemory w0, dp, 2
                            If w0 = 104 Then
                                w1 = 0
                                CopyMemory w1, dp + 2, 2
                                If w1 = 101 Then
                                    Dim wa As Long, wb As Long, wc As Long
                                    wa = 0: wb = 0: wc = 0
                                    CopyMemory wa, dp + 4, 2
                                    CopyMemory wb, dp + 6, 2
                                    CopyMemory wc, dp + 8, 2
                                    If wa = 108 And wb = 108 And wc = 111 Then tag = tag & "+helloPtr"
                                End If
                            End If
                        End If
                    End If
                    If Len(tag) > 0 Then
                        If nh > 0 Then hits = hits & ","
                        hits = hits & "{" & Q("addr") & ":" & CStr(p) & "," & Q("vt") & ":" & CStr(vt) & "," & Q("data") & ":" & CStr(dp) & "," & Q("tag") & ":" & Q(tag) & "," & Q("deltaFromFrame") & ":" & CStr(p - curFrame) & "}"
                        nh = nh + 1
                        If nh >= 16 Then Exit For
                    End If
                End If
            End If
        Next
    End If
    s = s & "," & Q("valueScan") & ":[" & hits & "]"
    s = s & "," & Q("valueScanCount") & ":" & CStr(nh)

    Dim cands As String, nc As Long
    Dim e As Long, rtmi As Long, oi As Long, pod As Long
    cands = ""
    nc = 0
    If curFrame > &H10000 Then
        For e = curFrame - &H600 To curFrame + &H600 Step 4
            If e > &H10000 And e < &H7FFFFFFF Then
                If IsBadReadPtr(e - &H50, 4) = 0 Then
                    rtmi = 0
                    CopyMemory rtmi, e - &H50, 4
                    If rtmi > &H10000 And rtmi < &H7FFFFFFF Then
                        If IsBadReadPtr(rtmi, 8) = 0 Then
                            oi = 0
                            CopyMemory oi, rtmi, 4
                            pod = 0
                            If oi > &H10000 And oi < &H7FFFFFFF Then
                                If IsBadReadPtr(oi + 4, 4) = 0 Then CopyMemory pod, oi + 4, 4
                            End If
                            If pod > &H10000 And pod < &H7FFFFFFF Then
                                If nc > 0 Then cands = cands & ","
                                cands = cands & "{" & Q("ebp") & ":" & CStr(e) _
                                    & "," & Q("rtmi") & ":" & CStr(rtmi) _
                                    & "," & Q("objInfo") & ":" & CStr(oi) _
                                    & "," & Q("pod") & ":" & CStr(pod) _
                                    & "," & Q("rtmiMatchesRef") & ":" & B(rtmi = rtmiRef) & "}"
                                nc = nc + 1
                                If nc >= 24 Then Exit For
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End If
    s = s & "," & Q("ebpCandidates") & ":[" & cands & "]"
    s = s & "," & Q("ebpCandidateCount") & ":" & CStr(nc)
    FrameStateJson = s & "}"
End Function

' The next finding:
' "evaluate inside the current frame" behaviour is a GLOBAL, not an argument --
' WORD[VBA6+0x17D076] bit8 (0x100). sub_FBB3D6C sets it before the worker call and
' restores it after; EbExecuteLine never touches it. So the caller can try the same
' thing: save the word, OR in 0x100, call, restore. If locals then resolve, G3 is
' closed without IExprSrv. This is a WRITE to engine state, so it is always
' save/restore and the original value is reported.
Public Function ExecLineJson(ByVal line As String, ByVal a2 As Long, ByVal a3 As Long, ByVal a4 As Long, ByVal setScope As Boolean, ByVal dumpCtx As Boolean, ByVal setFrame As Boolean) As String
    Dim hMod As Long, engineA As Long, engineB As Long, rc As Long
    Dim wBefore As Long, wSet As Long
    Dim buf(63) As Byte
    Dim ctxPtr As Long
    Dim fBefore As Long, fUse As Long
    Dim pHead As Long, node As Long, nNext As Long, nKind As Long
    Dim k As Long
    Dim s As String
    hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    s = s & Q("moduleBase") & ":" & CStr(hMod)
    If hMod = 0 Then
        ExecLineJson = FailSkip(s & "," & Q("called") & ":false," & Q("skipped") & ":" & Q("VBA6.DLL is not loaded"), "not_loaded", "VBA6.DLL is not loaded")
        Exit Function
    End If
    engineA = 0
    engineB = 0
    CopyMemory engineA, hMod + &H17D15C, 4
    CopyMemory engineB, hMod + &H17D0F4, 4
    s = s & "," & Q("engineAt17D15C") & ":" & CStr(engineA)
    s = s & "," & Q("engineAt17D0F4") & ":" & CStr(engineB)
    If engineA = 0 Then
        s = s & "," & Q("called") & ":false
        ExecLineJson = FailSkip(s & "," & Q("skipped") & ":" & Q("engine singleton at VBA6+0x17D15C is 0 (host has not initialised it); EbExecuteLine would dereference null"), "engine_zero", "engine singleton at VBA6+0x17D15C is 0 (host has not initialised it); EbExecuteLine would dereference null")
        Exit Function
    End If
    wBefore = 0
    CopyMemory wBefore, hMod + &H17D076, 2
    s = s & "," & Q("scopeWordBefore") & ":" & CStr(wBefore)
    s = s & "," & Q("inEvalWatchBit") & ":" & CStr(wBefore And &H100)
    ctxPtr = 0
    If dumpCtx Then
        buf(0) = 0
        ctxPtr = VarPtr(buf(0))
        a2 = ctxPtr
        s = s & "," & Q("ctxBuffer") & ":" & CStr(ctxPtr)
    End If
    If setScope Then
        wSet = wBefore Or &H100
        CopyMemoryToPtr hMod + &H17D076, wSet, 2
        s = s & "," & Q("scopeWordSet") & ":" & CStr(wSet)
    End If
    ' Walk engine+0x2F0, take the first kind==1 node's payload (node+8) and make
    ' it the current evaluation frame. This is what the IDE's own sub_FBC3163 does.
    fBefore = 0
    fUse = 0
    CopyMemory fBefore, hMod + &H17D0A4, 4
    s = s & "," & Q("evalFrameBefore") & ":" & CStr(fBefore)
    If setFrame Then
        pHead = 0
        If engineB <> 0 Then CopyMemory pHead, engineB + &H2F0, 4
        node = pHead
        For k = 0 To 31
            If node = 0 Then Exit For
            CopyMemory nKind, node + 4, 4
            If nKind = 1 Then
                CopyMemory fUse, node + 8, 4
                Exit For
            End If
            CopyMemory nNext, node, 4
            node = nNext
        Next
        s = s & "," & Q("frameNodeChosen") & ":" & CStr(node)
        s = s & "," & Q("frameValue") & ":" & CStr(fUse)
        If fUse <> 0 Then
            CopyMemoryToPtr hMod + &H17D0A4, fUse, 4
            s = s & "," & Q("evalFrameSet") & ":true"
        Else
            s = s & "," & Q("evalFrameSet") & ":false"
        End If
    End If
    rc = 0
    On Error Resume Next
    rc = EbExecuteLine(StrPtr(line), a2, a3, a4)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number) & "," & Q("vbaErrorDesc") & ":" & Q(Err.Description)
        Err.Clear
    End If
    On Error GoTo 0
    If setScope Then
        CopyMemoryToPtr hMod + &H17D076, wBefore, 2
        s = s & "," & Q("scopeRestored") & ":true"
    End If
    If setFrame Then
        CopyMemoryToPtr hMod + &H17D0A4, fBefore, 4
        s = s & "," & Q("evalFrameRestored") & ":true"
    End If
    s = s & "," & Q("called") & ":true"
    s = s & "," & Q("line") & ":" & Q(line)
    s = s & "," & Q("a2") & ":" & CStr(a2) & "," & Q("a3") & ":" & CStr(a3) & "," & Q("a4") & ":" & CStr(a4)
    s = s & "," & Q("ret") & ":" & CStr(rc)
    s = s & "," & Q("retHex") & ":" & Q("0x" & Hex$(rc))
    If dumpCtx Then
        Dim hexs As String
        Dim i As Long
        hexs = ""
        For i = 0 To 31
            hexs = hexs & Right$("0" & Hex$(buf(i)), 2)
        Next
        s = s & "," & Q("ctxDump32") & ":" & Q(hexs)
    End If
    ExecLineJson = s & "}"
End Function

Public Function ProbeRunJson(ByVal runBreakProbes As Boolean, ByVal force As Boolean, ByVal ident As String, ByVal step As String, ByVal frameIndex As Long) As String
    Dim mode As Long, stepMode As Long, evalWatch As Long
    Dim s As String, why As String
    Dim doBreak As Boolean
    Dim engineVal As Long, pcodeFlag As Long, hMod As Long, pGlobal As Long
    Dim idRes As Long, rcIdent As Long
    Dim pHead As Long, nodePtr As Long, nNext As Long, nKind As Long, nV8 As Long, nWC As Long, nameWC As String, k As Long, nodes As String

    mode = EbMode()
    stepMode = EbGetStepMode()
    evalWatch = EbIsInEvalWatch()

    s = "{"
    s = s & Q("ebMode") & ":" & CStr(mode)
    s = s & "," & Q("stepMode") & ":" & CStr(stepMode)
    s = s & "," & Q("inEvalWatch") & ":" & CStr(evalWatch)
    s = s & "," & ModuleInfoJson()
    s = s & ThreadInfoJson()

    ' An earlier batch: engine global -> the p-code / native branch flag at +0x2F4
    hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    If hMod <> 0 Then
        pGlobal = hMod + &H17D0F4
        CopyMemory engineVal, pGlobal, 4
        If engineVal <> 0 Then CopyMemory pcodeFlag, engineVal + &H2F4, 4
        s = s & "," & Q("engine") & ":" & CStr(engineVal)
        s = s & "," & Q("pcodeFlag") & ":" & CStr(pcodeFlag)

        ' An earlier batch: walk the node list and report each node's kind, so RE can tell
        ' which name branch the frames actually take
        If runBreakProbes Then
            CopyMemory pHead, engineVal + &H2F0, 4
            nodePtr = pHead
            For k = 0 To 15
                If nodePtr = 0 Then Exit For
                CopyMemory nNext, nodePtr, 4
                CopyMemory nKind, nodePtr + 4, 4
                CopyMemory nV8, nodePtr + 8, 4
                nWC = 0
                nameWC = ""
                CopyMemory nWC, nodePtr + &HC, 4
                If nWC <> 0 Then nameWC = ReadBstrAt(nWC)
                If Len(nodes) > 0 Then nodes = nodes & ","
                nodes = nodes & "{" & Q("i") & ":" & CStr(k) & "," & Q("node") & ":" & CStr(nodePtr) _
                    & "," & Q("kind") & ":" & CStr(nKind) & "," & Q("v8") & ":" & CStr(nV8) _
                    & "," & Q("wC") & ":" & CStr(nWC) & "," & Q("nameAtWC") & ":" & Q(nameWC) & "}"
                nodePtr = nNext
            Next
            s = s & "," & Q("nodes") & ":[" & nodes & "]"
        End If
    End If

    If Len(ident) > 0 Then
        idRes = 0
        rcIdent = EbIsValidIdent(StrPtr(ident), idRes)
        s = s & "," & Q("ident") & ":" & Q(ident)
        s = s & "," & Q("identRc") & ":" & CStr(rcIdent)
        s = s & "," & Q("identValid") & ":" & CStr(idRes)
    End If

    doBreak = False
    why = ""
    If Not runBreakProbes Then
        why = "not requested (pass run_break_probes=true)"
    ElseIf mode <> 2 And Not force Then
        why = "EbMode=" & CStr(mode) & " (need 2 = break context); pass force=true to try anyway"
    Else
        doBreak = True
    End If

    If Not doBreak Then
        s = s & "," & Q("breakProbes") & ":false"
        ProbeRunJson = FailSkip(s, "refused_by_design", why)
        Exit Function
    End If

    Dim rc As Long, cnt As Long, i As Long, names As String
    Dim loIdx As Long, hiIdx As Long
    Dim a4 As Long, a5 As Long, mname As String, pname As String
    Dim c1 As Long, c2 As Long, idx As Long, node As Long

    ' stepped mode: return as soon as the next call would be the risky one, so a
    ' crash cannot hide the values already collected
    If LCase$(step) = "count" Then
        Dim c0 As Long
        c0 = -1
        rc = EbGetCallstackCount(c0)
        s = s & "," & Q("breakProbes") & ":true"
        s = s & "," & Q("step") & ":""count"""
        s = s & "," & Q("callstackCountRc") & ":" & CStr(rc)
        s = s & "," & Q("callstackCount") & ":" & CStr(c0)
        ProbeRunJson = s & "}"
        Exit Function
    End If

    cnt = -1
    rc = EbGetCallstackCount(cnt)
    s = s & "," & Q("breakProbes") & ":true"
    s = s & "," & Q("callstackCountRc") & ":" & CStr(rc)
    s = s & "," & Q("callstackCount") & ":" & CStr(cnt)

    ' frame indices are 0-BASED: passing cnt crashed inside EbGetCallstackFunction
    ' (mov eax,[edi+4] with edi=NULL). Clamp so a bad index can never fault.
    If LCase$(step) = "frame" Then
        loIdx = frameIndex
        hiIdx = frameIndex
    Else
        loIdx = 0
        hiIdx = cnt - 1
    End If
    If loIdx < 0 Then loIdx = 0
    If hiIdx > cnt - 1 Then hiIdx = cnt - 1
    If cnt > 0 Then
        For i = loIdx To hiIdx
            mname = ""
            pname = ""
            a4 = 0
            a5 = 0
            rc = EbGetCallstackFunction(i, mname, pname, a4, a5)
            If Len(names) > 0 Then names = names & ","
            names = names & "{" & Q("frame") & ":" & CStr(i) & "," & Q("rc") & ":" & CStr(rc) _
                & "," & Q("module") & ":" & Q(Fix16(mname)) _
                & "," & Q("proc") & ":" & Q(Fix16(pname)) _
                & "," & Q("modBytes") & ":" & CStr(SysStringLen(StrPtr(mname))) _
                & "," & Q("procBytes") & ":" & CStr(SysStringLen(StrPtr(pname))) _
                & "," & Q("a4") & ":" & CStr(a4) & "," & Q("a5") & ":" & CStr(a5) & "}"
        Next
    End If
    s = s & "," & Q("frames") & ":[" & names & "]"

    c1 = 0
    c2 = 0
    rc = EbGetCurrentFunction(c1, c2)
    s = s & "," & Q("currentFunctionRc") & ":" & CStr(rc)
    s = s & "," & Q("currentFunction") & ":" & CStr(c1)
    s = s & "," & Q("currentFunction2") & ":" & CStr(c2)

    idx = -1
    ' R5: the RETURN VALUE is the selected callstack node pointer (0 = none);
    ' the out param is the 0-based index. It was never a status code.
    node = EbGetDebWndCallstackIndex(idx)
    s = s & "," & Q("debWndNode") & ":" & CStr(node)
    s = s & "," & Q("debWndIndex") & ":" & CStr(idx)
    s = s & "," & Q("fExitBreakMode") & ":" & CStr(EbFExitBreakMode())
    s = s & "}"
    ProbeRunJson = s
End Function

' ---------------------------------------------------------------------------
' Report which thread we are on.
'   tid          = the thread that is executing this probe
'   mainTid      = the thread that owns the VB6 IDE main window
'   onMainThread = true when they are the same thread
' If onMainThread is already true, the "must move the payload to the
' main thread" premise is ALREADY satisfied for this call path.
' ---------------------------------------------------------------------------
Private Function ThreadInfoJson() As String
    Dim tid As Long, pid As Long, hWnd As Long, wpid As Long, mtid As Long
    Dim s As String
    tid = GetCurrentThreadId()
    pid = GetCurrentProcessId()
    hWnd = FindWindowW(StrPtr("ThunderRT6Main"), 0)
    If hWnd = 0 Then hWnd = FindWindowW(0, StrPtr("Microsoft Visual Basic"))
    mtid = 0
    If hWnd <> 0 Then mtid = GetWindowThreadProcessId(hWnd, wpid)
    s = "," & Q("tid") & ":" & CStr(tid) & "," & Q("pid") & ":" & CStr(pid)
    s = s & "," & Q("mainWnd") & ":" & CStr(hWnd) & "," & Q("mainWndTid") & ":" & CStr(mtid)
    If mtid <> 0 Then
        s = s & "," & Q("onMainThread") & ":" & B(tid = mtid)
    Else
        s = s & "," & Q("onMainThread") & ":" & Q("unknown")
    End If
    ThreadInfoJson = s
End Function

' Parse a single argument token: decimal, or hex with an 0x prefix. Handles
' values above &H7FFFFFFF (as negative Longs) without raising overflow.
Private Function ParseArgToken(ByVal t As String) As Long
    Dim s As String
    s = Trim$(t)
    If Len(s) = 0 Then
        ParseArgToken = 0
    ElseIf Len(s) > 2 Then
        If LCase$(Left$(s, 2)) = "0x" Then
            ParseArgToken = HexToLong(Mid$(s, 3))
        Else
            ParseArgToken = CLng(s)
        End If
    Else
        ParseArgToken = CLng(s)
    End If
End Function

Private Function HexToLong(ByVal h As String) As Long
    Dim i As Long, c As Long
    Dim v As Double
    For i = 1 To Len(h)
        c = InStr("0123456789abcdef", LCase$(Mid$(h, i, 1))) - 1
        If c < 0 Then c = 0
        v = v * 16 + c
    Next
    If v > 2147483647# Then v = v - 4294967296#
    HexToLong = CLng(v)
End Function

' ---------------------------------------------------------------------------
' An N-argument caller.
' CreateIExprSrvObj is __stdcall with THREE arguments (retn 0Ch) -- the original
' CallRvaJson thunk only builds 0..2, which is exactly why the RE-specified call
' could never be made. This variant takes a comma-separated argument list and
' reuses the same proven mechanism:
'     mov ecx,[VBA6+engine]   (harmless for __stdcall, needed for __thiscall)
'     push aN ... push a1     (right-to-left == __stdcall layout)
'     mov eax,target ; call eax ; ret 16
' The trailing "ret 16" consumes the four arguments CallWindowProcA forwards,
' so the caller's stack stays balanced. Max 6 arguments.
' ---------------------------------------------------------------------------
Public Function CallRvaArgsJson(ByVal rva As Long, ByVal argsCsv As String, Optional ByVal dumpLen As Long = 0, Optional ByVal hModOverride As Long = 0) As String
    Dim hMod As Long, target As Long, mem As Long, rc As Long
    Dim b(95) As Byte
    Dim p As Long, i As Long, argc As Long
    Dim parts() As String
    Dim argv(0 To 5) As Long
    Dim s As String
    hMod = hModOverride
    If hMod = 0 Then hMod = GetModuleHandleW(StrPtr("VBA6.DLL"))
    s = "{"
    If hMod = 0 Then
        CallRvaArgsJson = FailErr(s & Q("error") & ":" & Q("module not loaded"), "not_loaded", "module not loaded")
        Exit Function
    End If
    target = hMod + rva
    parts = Split(argsCsv, ",")
    argc = UBound(parts) + 1
    If argc > 6 Then argc = 6
    If argc < 0 Then argc = 0
    For i = 0 To argc - 1
        argv(i) = ParseArgToken(parts(i))
    Next
    s = s & Q("rva") & ":" & Q("0x" & Hex$(rva)) & "," & Q("target") & ":" & CStr(target)
    s = s & "," & Q("argc") & ":" & CStr(argc) & "," & Q("args") & ":["
    For i = 0 To argc - 1
        If i > 0 Then s = s & ","
        s = s & CStr(argv(i))
    Next
    s = s & "]"
    mem = VirtualAlloc(0, 256, &H1000, &H40)
    If mem = 0 Then
        CallRvaArgsJson = FailErr(s & "," & Q("error") & ":" & Q("VirtualAlloc failed"), "alloc_failed", "VirtualAlloc failed")
        Exit Function
    End If
    p = 0
    If hModOverride = 0 Then b(p) = &H8B: b(p + 1) = &HD: Put32 b, p + 2, hMod + &H17D0F4: p = p + 6
    For i = argc - 1 To 0 Step -1
        b(p) = &H68: Put32 b, p + 1, argv(i): p = p + 5
    Next
    b(p) = &HB8: Put32 b, p + 1, target: p = p + 5
    b(p) = &HFF: b(p + 1) = &HD0: p = p + 2
    b(p) = &HC2: b(p + 1) = &H10: b(p + 2) = &H0: p = p + 3
    CopyMemoryToPtr mem, b(0), p
    rc = 0
    On Error Resume Next
    rc = CallWindowProcA(mem, 0, 0, 0, 0)
    If Err.Number <> 0 Then
        s = s & "," & Q("vbaError") & ":" & CStr(Err.Number)
        Err.Clear
    End If
    On Error GoTo 0
    s = s & "," & Q("ret") & ":" & CStr(rc) & "," & Q("retHex") & ":" & Q("0x" & Hex$(rc))
    ' ATOMIC create+dump: dumping the returned pointer in a SEPARATE request lets the
    ' block be recycled first (measured: three different pointers returned byte-identical
    ' content that was in fact our own HTTP response). Dump it here, immediately.
    If dumpLen > 0 Then
        If dumpLen > 512 Then dumpLen = 512
        dumpLen = (dumpLen \ 4) * 4
        If dumpLen < 4 Then dumpLen = 4
        s = s & "," & Q("dumpLen") & ":" & CStr(dumpLen)
        If rc < &H10000 Or rc > &H7FFFFFFF Then
            s = s & "," & Q("dumpError") & ":" & Q("return value is not a pointer")
        ElseIf IsBadReadPtr(rc, dumpLen) <> 0 Then
            s = s & "," & Q("dumpError") & ":" & Q("return value is not readable")
        Else
            Dim db() As Byte
            ReDim db(dumpLen - 1)
            CopyMemory db(0), rc, dumpLen
            s = s & "," & Q("dumpHex") & ":" & Q(BytesToHex(db, dumpLen))
        End If
    End If
    CallRvaArgsJson = s & "}"
End Function

Private Function BytesToHex(ByRef b() As Byte, ByVal n As Long) As String
    Dim i As Long, s As String
    s = ""
    For i = 0 To n - 1
        s = s & Right$("0" & Hex$(b(i)), 2)
    Next
    BytesToHex = s
End Function

' =====================================================================
' RE-1 (34a): STRUCTURED ERROR RECORD (read-only, zero engine calls)
' =====================================================================
' The VBA6 exception record is a 32-byte global (RE rounds 117 / 60) and the G2
' runtime probe CONFIRMED its field offsets across TWO IDE processes
' (fix/TASK_RUNTIME_PROBES/evidence/g2_mapping.json: verdict=CONFIRMED,
' eventVerdicts=[CONFIRMED,CONFIRMED], desc +0x08 / src +0x04 / num +0x1C).
'
' WHY A MEMORY READ AND NOT AN EXPORT: while the modal runtime-error box owns the
' IDE thread this code path must make NO VBA6 call at all (measured discipline
' documented in CodeTools.ToolLastError). Reading two globals plus two BSTRs with
' CopyMemory makes no call and does not touch engine state. EbExceptionHandled is
' NEVER called: it CLEARS the exception state bit (dword_FC0D010 = 0) and would
' blind the IDE's own error handling.
'
' RECORD SELECTOR: static disassembly says the record is chosen by
' dword_FC0D074 AND &H1000000 (0 -> record A at +0x17D078, non-zero -> record B
' at +0x17D0A8). The G2 probe read 0x17D010 instead, which is the exception STATE
' bit, not the selector; both agreed (0) in the captured event, so the confirmed
' field mapping stays valid. This reader uses the static selector and FALLS BACK
' to the other record when the selected one reads empty, publishing which path it
' took (selectorBasis = dword074_bit24 | fallback_probe).
'
' Returns a Dictionary (never raises). Keys: readOk, recordUsed, raw010,
' dword074, bit24, chosen, selectorBasis, source, description, hresult, number,
' numberBasis, sourceOk/sourceChars/sourceBytes/sourceError, description<same>,
' why. Nothing is invented: unreadable -> Null + why.
Public Function ErrRecordInfo() As Object
    Dim d As Object
    Dim hMod As Long, selAddr As Long, recA As Long, recB As Long, stateAddr As Long
    Dim state As Long, selDword As Long, bit24 As Boolean
    Dim used As String, basis As String, recAddr As Long, hr As Long
    Dim src As String, desc As String
    On Error Resume Next
    Set d = CreateObject("Scripting.Dictionary")
    d("readOk") = False
    d("recordUsed") = Null
    d("raw010") = Null
    d("dword074") = Null
    d("bit24") = Null
    d("chosen") = Null
    d("selectorBasis") = "unavailable"
    d("source") = Null
    d("description") = Null
    d("hresult") = Null
    d("number") = Null
    d("numberBasis") = "unavailable"
    d("sourceOk") = False
    d("descriptionOk") = False
    d("sourceChars") = 0
    d("sourceBytes") = 0
    d("descriptionChars") = 0
    d("descriptionBytes") = 0
    d("sourceError") = Null
    d("descriptionError") = Null
    d("why") = Null
    hMod = Vba6Base()
    If hMod = 0 Then
        d("why") = "vba6_not_loaded"
        Set ErrRecordInfo = d
        Exit Function
    End If
    selAddr = hMod + &H17D074
    recA = hMod + &H17D078
    recB = hMod + &H17D0A8
    stateAddr = hMod + &H17D010
    If IsBadReadPtr(selAddr, 4) <> 0 Then
        d("why") = "selector_unreadable"
        Set ErrRecordInfo = d
        Exit Function
    End If
    state = 0
    CopyMemory state, stateAddr, 4
    selDword = 0
    CopyMemory selDword, selAddr, 4
    bit24 = ((selDword And &H1000000&) <> 0)
    d("raw010") = state
    d("dword074") = selDword
    d("bit24") = bit24
    ' 34a LIVE MEASUREMENT: record A is the active record BOTH while the modal box is up
    ' (bit24 = 1) and after it settles (bit24 = 0), so bit24 tracks the MODAL DIALOG STATE
    ' and is NOT a record selector. A is therefore read first, and B is only an
    ' empty-record fallback probe (its basis stays "fallback_probe" when it is used).
    used = "A"
    basis = "unavailable"
    recAddr = recA
    If IsBadReadPtr(recA, &H20) <> 0 Then
        used = "B"
        basis = "fallback_probe"
        recAddr = recB
        If IsBadReadPtr(recB, &H20) <> 0 Then
            d("chosen") = used
            d("selectorBasis") = basis
            d("why") = "record_unreadable"
            Set ErrRecordInfo = d
            Exit Function
        End If
    End If
    hr = 0
    CopyMemory hr, recAddr + &H1C, 4
    src = ErrReadBstr(recAddr + &H4, d, "source")
    desc = ErrReadBstr(recAddr + &H8, d, "description")
    If (Not CBool(d("sourceOk"))) And (Not CBool(d("descriptionOk"))) And hr = 0 Then
        ' the chosen record is EMPTY: probe the other one (never used to paper over a
        ' readable-but-failing read)
        If used = "A" Then
            recAddr = recB
            used = "B"
        Else
            recAddr = recA
            used = "A"
        End If
        basis = "fallback_probe"
        If IsBadReadPtr(recAddr, &H20) = 0 Then
            d("sourceError") = Null
            d("descriptionError") = Null
            hr = 0
            CopyMemory hr, recAddr + &H1C, 4
            src = ErrReadBstr(recAddr + &H4, d, "source")
            desc = ErrReadBstr(recAddr + &H8, d, "description")
        End If
    End If
    d("chosen") = used
    d("selectorBasis") = basis
    d("recordUsed") = used
    If CBool(d("sourceOk")) Then d("source") = src
    If CBool(d("descriptionOk")) Then d("description") = desc
    If (Not CBool(d("sourceOk"))) And (Not CBool(d("descriptionOk"))) And hr = 0 Then
        d("why") = "record_empty"
        Set ErrRecordInfo = d
        Exit Function
    End If
    d("hresult") = hr
    If (hr And &H1FFF0000&) = &HA0000& Then
        d("number") = (hr And &HFFFF&)
        d("numberBasis") = "hresult_low16"
    Else
        d("numberBasis") = "facility_not_vb"
    End If
    d("readOk") = True
    Set ErrRecordInfo = d
End Function

' Read one BSTR out of the exception record. Bounded, guarded, never raises:
' a bad pointer / bad length leaves the field Null and records the reason.
Private Function ErrReadBstr(ByVal fieldAddr As Long, ByRef d As Object, ByVal which As String) As String
    Dim n As Long, i As Long, s As String, w As Long, p As Long
    Dim okKey As String, charsKey As String, bytesKey As String, errKey As String
    okKey = which & "Ok"
    charsKey = which & "Chars"
    bytesKey = which & "Bytes"
    errKey = which & "Error"
    ErrReadBstr = ""
    d(okKey) = False
    d(charsKey) = 0
    d(bytesKey) = 0
    If fieldAddr = 0 Then
        d(errKey) = "null_pointer"
        Exit Function
    End If
    If IsBadReadPtr(fieldAddr, 4) <> 0 Then
        d(errKey) = "pointer_unreadable"
        Exit Function
    End If
    ' the record stores a POINTER to the BSTR data; fieldAddr is the address of that slot
    p = 0
    CopyMemory p, fieldAddr, 4
    If p = 0 Then
        d(errKey) = "null_pointer"
        Exit Function
    End If
    If IsBadReadPtr(p, 8) <> 0 Then
        d(errKey) = "pointer_unreadable"
        Exit Function
    End If
    n = SysStringLen(p)
    If n <= 0 Then
        d(errKey) = "empty"
        Exit Function
    End If
    If n > 4096 Then
        d(errKey) = "length_out_of_range"
        Exit Function
    End If
    If IsBadReadPtr(p, n * 2 + 2) <> 0 Then
        d(errKey) = "payload_unreadable"
        Exit Function
    End If
    For i = 0 To n - 1
        w = 0
        CopyMemory w, p + i * 2, 2
        s = s & ChrW$(w And &HFFFF&)
    Next
    d(okKey) = True
    d(charsKey) = n
    d(bytesKey) = n * 2
    ErrReadBstr = s
End Function

Public Function IdeEbMode() As Long
    ' (i18n): expose the VBA6 debug-engine mode so IDE-state
    ' detection can stop reading the localized main-window caption. EbMode is a
    ' language-independent signal: 2 = break context, 0 = no debug context.
    On Error Resume Next
    IdeEbMode = EbMode()
End Function
