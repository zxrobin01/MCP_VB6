Attribute VB_Name = "TraceQuery"
Option Explicit

' ============================================================================
' trace_query -- structured, filterable, resumable READER for the McpTrace
' JSONL file sink (the read side of inject_trace / logpoint).  READ-ONLY.
'
' CONTRACT AUTHORITY: tools/trace_query.py (the "A form") and
' docs/design/DEBUG_V1170_BATCH_SPEC.md section 2 plus the BATCH-O1..O10
' rulings in docs/design/DEBUG_LEAD_RULINGS_20261028.md.  The receipt KEY SET
' and the two CLOSED sets (8 failure reasons, 7 warning codes) are the A
' form's, verbatim.  An argument outside its declared domain RAISES a VB error
' (dispatch error envelope, ruling BATCH-O6) and is NEVER one of the 8
' observation failures.
'
' B-ONLY differences (all documented and D-numbered in the spec):
'   D2  path is optional: explicit path, else MCP_TRACE_FILE, else the TEMP
'       directory plus mcp_trace.jsonl; pathFrom says which one was used and
'       is null when none resolves (then file_not_found).  The sink is never
'       created (S6).
'   D7  only TOP-LEVEL duplicate JSON keys are detected (the sink writes flat
'       objects; nested objects only arrive if an external tool rewrites it).
'   D11 a serialized receipt (or a jsonl raw body) above about 1 MiB is cut
'       down with truncatedBy="bytes_budget" and cursor.nextOffset is REWOUND
'       to the first unreturned event, so a resume loses nothing.  Why: the
'       transport silently DROPS a frame above 8 MiB (PipeTransport
'       NP_MAX_FRAME), so returning less is the only honest option.
'   D12 fileInfo.mtime is the REAL UTC last-write time from GetFileTime +
'       FileTimeToSystemTime (never local time, never invented).
'   D16 format="jsonl" returns the raw matched lines only (no envelope); the
'       DISPATCH layer must bypass the receipt wrapper for that body (see the
'       "trace_query" case in CodeTools.CallToolInner).  count_only=true wins
'       over format, because an empty body would be indistinguishable from a
'       transport failure (ruling BATCH-O8).
'
' ZERO-WRITE (S1): this module never creates, truncates, appends to, deletes
' or re-times a file, and it never creates the sink or a directory.  It
' deliberately does NOT use VB "Open ... For Binary Access Read": measured
' VB6 behaviour (documented in Encoding.bas) is that this statement CREATES a
' missing file, which would violate S1.  The file is opened with
' CreateFileW(OPEN_EXISTING, GENERIC_READ) and read with ReadFile; the share
' mode is READ|WRITE|DELETE so a running sink can keep appending.
'
' NO VBIDE ACCESS (S2/S7): nothing here touches the IDE, the project, the
' debugger or the execution state; it only reads bytes and file metadata.
' No cross-call cache: every call re-derives everything (S10).
'
' The scan consumes COMPLETE lines only; a trailing half line is never parsed
' and never consumed (cursor.partialLineBytes reports it), and cursor
' nextOffset always points at the next byte to read.  events[].seq is
' CALL-LOCAL (a 1-based physical line number inside this window), never a
' cursor: the ONLY persistent cursor is since_offset (ruling A.1.6).
' ============================================================================

' ---- identity / budgets (identical to tools/trace_query.py) ----
Private Const TQ_TOOL As String = "trace_query"
Private Const TQ_VERSION As String = "1.0"
Private Const TQ_DEFAULT_LIMIT As Long = 100
Private Const TQ_MAX_LIMIT As Long = 1000
Private Const TQ_MIN_BYTES As Long = 4096
Private Const TQ_DEFAULT_MAX_BYTES As Long = 1048576
Private Const TQ_MAX_BYTES As Long = 8388608
Private Const TQ_MAX_FILE_BYTES As Double = 268435456#
Private Const TQ_HEAD_HASH_BYTES As Long = 65536
Private Const TQ_TRACE_PROBE_LINES As Long = 64
Private Const TQ_TRACE_PROBE_MAX_BYTES As Long = 1048576
Private Const TQ_MAX_BAD_LINES As Long = 50
Private Const TQ_MAX_DUP_LINES As Long = 10
Private Const TQ_EXCERPT_MAX As Long = 200
Private Const TQ_RECEIPT_BUDGET As Long = 1048576
Private Const TQ_MAX_JSON_DEPTH As Long = 64
Private Const TQ_ALIGN_SCAN_MAX As Long = 8388608

' ---- file access constants (Win32) ----
Private Const TQ_FILE_READ_DATA As Long = &H1&
Private Const TQ_FILE_READ_ATTRIBUTES As Long = &H80&
Private Const TQ_SHARE_ALL As Long = 7&
Private Const TQ_OPEN_EXISTING As Long = 3
Private Const TQ_INVALID_HANDLE As Long = -1&
Private Const TQ_FILE_BEGIN As Long = 0&
Private Const TQ_ERR_HANDLE_EOF As Long = 38&
Private Const TQ_ERR_FILE_NOT_FOUND As Long = 2&
Private Const TQ_ERR_PATH_NOT_FOUND As Long = 3&

' ---- failure reasons: a CLOSED, mutually exclusive set (8 values) ----
Private Const TQ_R_FILE_NOT_FOUND As String = "file_not_found"
Private Const TQ_R_REFUSED_PATH As String = "refused_path"
Private Const TQ_R_TOO_LARGE As String = "too_large"
Private Const TQ_R_IO_ERROR As String = "io_error"
Private Const TQ_R_BAD_CURSOR As String = "bad_cursor"
Private Const TQ_R_BAD_JSONL As String = "bad_jsonl_line"
Private Const TQ_R_NOT_TRACE As String = "not_a_trace_file"
Private Const TQ_R_ENCODING As String = "encoding_unsupported"

' ---- non-fatal warnings: a CLOSED set (7 values) ----
Private Const TQ_W_FILE_SHRANK As String = "file_shrank"
Private Const TQ_W_ALIGNED As String = "cursor_aligned_forward"
Private Const TQ_W_BAD_SKIPPED As String = "bad_lines_skipped"
Private Const TQ_W_PARTIAL As String = "partial_tail"
Private Const TQ_W_DUP_KEYS As String = "duplicate_keys"
Private Const TQ_W_FIELDS_MISSING As String = "fields_missing"
Private Const TQ_W_HEAD_ABSENT As String = "head_sha_absent"

' ---- JSON value kinds used by the lexical validator ----
Private Const TQ_K_OBJECT As Long = 1
Private Const TQ_K_ARRAY As Long = 2
Private Const TQ_K_STRING As Long = 3
Private Const TQ_K_INT As Long = 4
Private Const TQ_K_NUM As Long = 5
Private Const TQ_K_LITERAL As Long = 6

' ---- Win32 types (declared before the imports that use them) ----
Private Type TqFileTime
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type TqSystemTime
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

' ---- Win32 imports.  Plain names (no Alias) keep the project's bare-API
'      ratchet satisfied: each name is declared right here. ----
Private Declare Function CreateFileW Lib "kernel32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, ByRef lpCreationTime As TqFileTime, ByRef lpLastAccessTime As TqFileTime, ByRef lpLastWriteTime As TqFileTime) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (ByRef lpFileTime As TqFileTime, ByRef lpSystemTime As TqSystemTime) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

' ============================================================================
' PUBLIC ENTRY: the tool body.
'   args     the parsed argument object (a Scripting.Dictionary)
'   rawOut   TRUE only for a successful format="jsonl" answer: the caller must
'            put the returned text on the wire WITHOUT the receipt envelope
'   audit    a SHORT summary for the audit log (never raw trace content: S12)
' ============================================================================
Public Function ToolTraceQuery(ByVal args As Object, ByRef rawOut As Boolean, ByRef audit As String) As String
    Dim started As Long
    Dim fmtJsonl As Boolean
    Dim countOnly As Boolean
    Dim receipt As Object
    Dim okFlag As Boolean
    Dim rawText As String
    Dim jsonText As String
    Dim tqEvs As Object
    started = GetTickCount()
    rawOut = False
    audit = "ok=false"
    Set receipt = TqQuery(args, fmtJsonl, countOnly, started)
    okFlag = CBool(receipt("ok"))
    ' B-only safety layer (D11): cut the body down BEFORE it can hit the
    ' transport's silent frame drop.
    If okFlag Then TqApplyBudget receipt, fmtJsonl
    audit = TqAuditSummary(receipt, fmtJsonl)
    If okFlag And fmtJsonl And (Not countOnly) Then
        Set tqEvs = receipt("events")
        rawText = TqRawLines(tqEvs)
        If Len(rawText) <= TQ_RECEIPT_BUDGET Then
            rawOut = True
            ToolTraceQuery = rawText
            Exit Function
        End If
        ' Raw form too large: answer the structured receipt instead.  Returning a
        ' silently shortened raw body would be a fake success.
        TqStripRaw receipt
    End If
    jsonText = TqReceiptJson(receipt, started)
    ToolTraceQuery = jsonText
End Function

' ============================================================================
' receipt construction helpers
' ============================================================================
Private Function TqBaseReceipt(ByVal pathV As Variant, ByVal pathFromV As Variant, ByVal sinceV As Variant, ByVal filters As Object, ByVal strictVal As Boolean) As Object
    Dim d As Object
    Dim fi As Object
    Dim r As Object
    Dim cu As Object
    Dim cn As Object
    Dim f As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("ok") = False
    d("tool") = TQ_TOOL
    d("version") = TQ_VERSION
    d("pathFrom") = pathFromV
    Set fi = CreateObject("Scripting.Dictionary")
    fi("path") = pathV
    fi("exists") = Null
    fi("size") = Null
    fi("mtime") = Null
    fi("sha256Head64KiB") = Null
    fi("headBytesRead") = 0&
    Set d("fileInfo") = fi
    Set r = CreateObject("Scripting.Dictionary")
    r("startOffset") = Null
    r("endOffset") = Null
    r("bytesRead") = 0&
    r("sha256") = Null
    Set d("range") = r
    Set cu = CreateObject("Scripting.Dictionary")
    cu("sinceOffset") = sinceV
    cu("nextOffset") = Null
    cu("partialLineBytes") = 0&
    cu("reachedEof") = False
    Set d("cursor") = cu
    Set cn = CreateObject("Scripting.Dictionary")
    cn("linesScanned") = 0&
    cn("eventsMatched") = 0&
    cn("eventsSkipped") = 0&
    cn("badLines") = 0&
    cn("warningsTotal") = 0&
    cn("hits") = 0&
    Set d("counts") = cn
    d("hitsScope") = "window"
    Set d("events") = New Collection
    Set d("warnings") = New Collection
    Set f = CreateObject("Scripting.Dictionary")
    f("id") = filters("id")
    f("idPrefix") = filters("idPrefix")
    f("proc") = filters("proc")
    f("pid") = filters("pid")
    Set f("includeFields") = filters("includeFields")
    Set d("filters") = f
    d("truncated") = False
    d("truncatedBy") = Null
    Set d("badLines") = New Collection
    d("elapsedMs") = 0&
    d("strict") = strictVal
    Set TqBaseReceipt = d
End Function

' Fill a failure receipt: the closed-set reason plus the A form's extras.
Private Function TqFail(ByVal receipt As Object, ByVal reason As String, ByVal extras As Object) As Object
    Dim k As Variant
    Dim o As Object
    receipt("ok") = False
    receipt("reason") = reason
    If Not extras Is Nothing Then
        For Each k In extras.Keys
            If IsObject(extras(k)) Then
                Set o = extras(k)
                Set receipt(CStr(k)) = o
            Else
                receipt(CStr(k)) = extras(k)
            End If
        Next k
    End If
    Set receipt("events") = New Collection
    receipt("truncated") = False
    receipt("truncatedBy") = Null
    Set TqFail = receipt
End Function

' The A form's _finish(): warningsTotal + the window-scoped hits echo + elapsedMs.
Private Function TqFinish(ByVal receipt As Object, ByVal started As Long) As Object
    Dim c As Object
    Dim w As Object
    Set c = receipt("counts")
    Set w = receipt("warnings")
    c("warningsTotal") = w.Count
    receipt("hits") = c("hits")
    receipt("elapsedMs") = GetTickCount() - started
    Set TqFinish = receipt
End Function

' Serialize the final receipt (called AFTER the byte budget may have edited it,
' so the counters are recomputed here).
Private Function TqReceiptJson(ByVal receipt As Object, ByVal started As Long) As String
    Dim c As Object
    Dim w As Object
    Set c = receipt("counts")
    Set w = receipt("warnings")
    c("warningsTotal") = w.Count
    receipt("hits") = c("hits")
    receipt("elapsedMs") = GetTickCount() - started
    TqReceiptJson = JsonLib.JsonStringify(receipt)
End Function

' S12: the audit record carries only a summary, never trace content.
Private Function TqAuditSummary(ByVal receipt As Object, ByVal fmtJsonl As Boolean) As String
    Dim s As String
    Dim r As Object
    s = "ok=" & LCase$(CStr(receipt("ok")))
    s = s & " hits=" & CStr(receipt("hits"))
    If receipt.Exists("reason") Then s = s & " reason=" & CStr(receipt("reason"))
    Set r = receipt("range")
    s = s & " bytesRead=" & CStr(r("bytesRead"))
    If fmtJsonl Then s = s & " format=jsonl (raw lines are not logged)"
    TqAuditSummary = s
End Function

' ---------------------------------------------------------------------------
' B-only byte budget (BATCH-O3 / D11).  Takes the events collection down to the
' largest prefix whose serialized body fits TQ_RECEIPT_BUDGET, then REWINDS
' cursor.nextOffset to the first dropped event so nothing is silently lost.
' Returns TRUE when the receipt was cut.
' ---------------------------------------------------------------------------
Private Function TqApplyBudget(ByVal receipt As Object, ByVal rawMode As Boolean) As Boolean
    Dim evs As Object
    Dim allEv As Object
    Dim n As Long
    Dim lo As Long
    Dim hi As Long
    Dim midK As Long
    Dim cu As Object
    Dim cn As Object
    Dim dropped As Object
    TqApplyBudget = False
    If Not CBool(receipt("ok")) Then Exit Function
    Set evs = receipt("events")
    n = evs.Count
    If n <= 0 Then Exit Function
    If TqBodySize(receipt, rawMode) <= TQ_RECEIPT_BUDGET Then Exit Function
    Set allEv = evs
    lo = 0
    hi = n
    Do While lo < hi
        midK = CLng(Int((lo + hi + 1) / 2))
        Set receipt("events") = TqFirstN(allEv, midK)
        If TqBodySize(receipt, rawMode) <= TQ_RECEIPT_BUDGET Then
            lo = midK
        Else
            hi = midK - 1
        End If
    Loop
    Set receipt("events") = TqFirstN(allEv, lo)
    If lo >= n Then Exit Function
    Set dropped = allEv(lo + 1)
    Set cu = receipt("cursor")
    cu("nextOffset") = dropped("lineOffset")
    cu("partialLineBytes") = 0&
    cu("reachedEof") = False
    receipt("truncated") = True
    receipt("truncatedBy") = "bytes_budget"
    Set cn = receipt("counts")
    cn("hits") = lo
    TqApplyBudget = True
End Function

' Size metric for the budget: the structured JSON body, and for the jsonl form
' also the raw body plus a small margin (whichever is larger).
Private Function TqBodySize(ByVal receipt As Object, ByVal rawMode As Boolean) As Long
    Dim s As String
    Dim n As Long
    Dim evs As Object
    s = JsonLib.JsonStringify(receipt)
    n = Len(s)
    If rawMode Then
        Set evs = receipt("events")
        If TqRawLen(evs) + 1024 > n Then n = TqRawLen(evs) + 1024
    End If
    TqBodySize = n
End Function

Private Function TqRawLen(ByVal evs As Object) As Long
    Dim i As Long
    Dim ev As Object
    Dim n As Long
    n = 0
    For i = 1 To evs.Count
        Set ev = evs(i)
        n = n + Len(CStr(ev("raw"))) + 1
    Next i
    TqRawLen = n
End Function

' The jsonl body: every matched raw line (the decoded JSONL text, CR stripped)
' followed by one LF, exactly like the A form's stdout.
Private Function TqRawLines(ByVal evs As Object) As String
    Dim i As Long
    Dim ev As Object
    Dim s As String
    s = ""
    For i = 1 To evs.Count
        Set ev = evs(i)
        s = s & CStr(ev("raw")) & vbLf
    Next i
    TqRawLines = s
End Function

' Remove the internal raw key before a jsonl body falls back to the structured
' receipt: the documented events[] key set has exactly six keys.
Private Function TqStripRaw(ByVal receipt As Object) As Object
    Dim evs As Object
    Dim i As Long
    Dim ev As Object
    Set evs = receipt("events")
    For i = 1 To evs.Count
        Set ev = evs(i)
        If ev.Exists("raw") Then ev.Remove "raw"
    Next i
    Set TqStripRaw = receipt
End Function

Private Function TqFirstN(ByVal evs As Object, ByVal k As Long) As Object
    Dim out As Object
    Dim i As Long
    Set out = New Collection
    If k > evs.Count Then k = evs.Count
    For i = 1 To k
        out.Add evs(i)
    Next i
    Set TqFirstN = out
End Function

' ============================================================================
' argument domain helpers.  A violation RAISES (dispatch envelope); it is never
' one of the 8 observation reasons (ruling BATCH-O6).
' ============================================================================
Private Sub TqRaiseParam(ByVal msg As String)
    Err.Raise vbObjectError + 9101, "TraceQuery", msg
End Sub

Private Function TqHasVal(ByVal d As Object, ByVal key As String) As Boolean
    TqHasVal = False
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    TqHasVal = True
End Function

Private Function TqArgStr(ByVal d As Object, ByVal key As String, ByRef outVal As String) As Boolean
    TqArgStr = False
    outVal = ""
    If Not TqHasVal(d, key) Then Exit Function
    If IsObject(d(key)) Then TqRaiseParam "argument '" & key & "' must be a string"
    If VarType(d(key)) <> vbString Then TqRaiseParam "argument '" & key & "' must be a string"
    outVal = CStr(d(key))
    TqArgStr = True
End Function

Private Function TqArgLong(ByVal d As Object, ByVal key As String, ByRef outVal As Long) As Boolean
    Dim v As Variant
    Dim vt As Integer
    Dim dd As Double
    TqArgLong = False
    outVal = 0
    If Not TqHasVal(d, key) Then Exit Function
    If IsObject(d(key)) Then TqRaiseParam "argument '" & key & "' must be an integer"
    vt = VarType(d(key))
    If vt = vbBoolean Or vt = vbString Then TqRaiseParam "argument '" & key & "' must be an integer"
    If vt <> vbDouble And vt <> vbLong And vt <> vbInteger And vt <> vbByte And vt <> vbCurrency And vt <> vbSingle Then
        TqRaiseParam "argument '" & key & "' must be an integer"
    End If
    v = d(key)
    dd = CDbl(v)
    If dd <> Int(dd) Then TqRaiseParam "argument '" & key & "' must be an integer"
    If dd > 2147483647# Then TqRaiseParam "argument '" & key & "' is outside the VB6 Long range"
    If dd < -2147483648# Then TqRaiseParam "argument '" & key & "' is outside the VB6 Long range"
    outVal = CLng(dd)
    TqArgLong = True
End Function

' include_fields: an array of strings OR one comma separated string (the A form
' accepts both); each list element is ALSO comma split, exactly like A.
' Returns Nothing when absent, null, or empty after normalisation.
Private Function TqArgFields(ByVal d As Object, ByVal key As String) As Object
    Dim out As Object
    Dim parts As Object
    Dim col As Object
    Dim item As Variant
    Dim v As Variant
    Dim raw As String
    Dim segs As Variant
    Dim i As Long
    Dim name As String
    Set out = Nothing
    If d Is Nothing Then Exit Function
    If Not d.Exists(key) Then Exit Function
    If IsNull(d(key)) Then Exit Function
    Set parts = New Collection
    If IsObject(d(key)) Then
        If TypeName(d(key)) <> "Collection" Then TqRaiseParam "argument '" & key & "' must be an array of strings or one comma separated string"
        Set col = d(key)
        For Each item In col
            If IsObject(item) Then TqRaiseParam "argument '" & key & "' must contain strings"
            If VarType(item) <> vbString Then TqRaiseParam "argument '" & key & "' must contain strings"
            parts.Add CStr(item)
        Next item
    Else
        If VarType(d(key)) <> vbString Then TqRaiseParam "argument '" & key & "' must be an array of strings or one comma separated string"
        parts.Add CStr(d(key))
    End If
    Set out = CreateObject("Scripting.Dictionary")
    For Each v In parts
        raw = CStr(v)
        segs = Split(raw, ",")
        For i = LBound(segs) To UBound(segs)
            name = Trim$(CStr(segs(i)))
            If Len(name) > 0 Then
                If Not out.Exists(name) Then out(name) = True
            End If
        Next i
    Next v
    If out.Count = 0 Then Set out = Nothing
    Set TqArgFields = out
End Function

' The filters includeFields echo, in declaration order, ALWAYS an array.
Private Function TqFieldList(ByVal fieldDict As Object) As Object
    Dim out As Object
    Dim k As Variant
    Set out = New Collection
    If fieldDict Is Nothing Then
        Set TqFieldList = out
        Exit Function
    End If
    For Each k In fieldDict.Keys
        out.Add CStr(k)
    Next k
    Set TqFieldList = out
End Function

' ============================================================================
' the query itself: the A form's ordered checks, then the scan and the receipt.
' ============================================================================
Private Function TqQuery(ByVal args As Object, ByRef fmtJsonl As Boolean, ByRef countOnly As Boolean, ByVal started As Long) As Object
    Dim receipt As Object
    Dim filters As Object
    Dim fieldFilter As Object
    Dim fl As Object
    Dim hasPath As Boolean
    Dim pathArg As String
    Dim hasId As Boolean
    Dim idArg As String
    Dim hasPrefix As Boolean
    Dim prefixArg As String
    Dim hasProc As Boolean
    Dim procArg As String
    Dim effectivePrefix As Variant
    Dim hasPid As Boolean
    Dim pidArg As Long
    Dim hasSince As Boolean
    Dim sinceArg As Long
    Dim hasLimit As Boolean
    Dim limitArg As Long
    Dim hasMax As Boolean
    Dim maxArg As Long
    Dim hasFmt As Boolean
    Dim fmtArg As String
    Dim hasExpect As Boolean
    Dim expectArg As String
    Dim strictVal As Boolean
    Dim sinkPath As Variant
    Dim pathFrom As Variant
    Dim sinkStr As String
    Dim envT As String
    Dim tempD As String
    Dim policy As String
    Dim failExtra As Object
    Dim cg As Object
    Dim ffe As Object
    Dim eps As Object
    Dim sb As Object
    Dim hFile As Long
    Dim openErr As Long
    Dim readErr As Boolean
    Dim fi As Object
    Dim sizeD As Double
    Dim mtimeV As Variant
    Dim wantLen As Long
    Dim headBuf() As Byte
    Dim headLen As Long
    Dim headSha As String
    Dim bomLen As Long
    Dim probe As Object
    Dim expectStr As String
    Dim effectiveStart As Long
    Dim shrank As Boolean
    Dim prevBuf() As Byte
    Dim prevLen As Long
    Dim alignedTo As Variant
    Dim remainD As Double
    Dim windowD As Double
    Dim windowLen As Long
    Dim buf() As Byte
    Dim bufLen As Long
    Dim readHitEof As Boolean
    Dim scan As Object
    Dim consumed As Long
    Dim endOffset As Long
    Dim r As Object
    Dim cu As Object
    Dim cn As Object
    Dim evs As Object
    Dim bads As Object
    Dim warnings As Object
    Dim truncatedBy As Variant
    Dim partialBytes As Long
    Dim reachedEof As Boolean
    Dim firstMatchOffset As Long
    Dim detail As String
    Dim fsSeen As Object
    fmtJsonl = False
    countOnly = False
    ' ---------- 1. argument domain (raises; before ANY filesystem access) ----------
    hasPath = TqArgStr(args, "path", pathArg)
    hasId = TqArgStr(args, "id", idArg)
    hasPrefix = TqArgStr(args, "id_prefix", prefixArg)
    hasProc = TqArgStr(args, "proc", procArg)
    If hasProc And hasPrefix Then TqRaiseParam "proc and id_prefix are mutually exclusive (proc is shorthand for id_prefix=proc plus a dot)"
    hasPid = TqArgLong(args, "pid", pidArg)
    hasSince = TqArgLong(args, "since_offset", sinceArg)
    hasLimit = TqArgLong(args, "limit", limitArg)
    If Not hasLimit Then limitArg = TQ_DEFAULT_LIMIT
    If limitArg < 1 Or limitArg > TQ_MAX_LIMIT Then TqRaiseParam "limit must be an integer in 1..1000"
    hasMax = TqArgLong(args, "max_bytes", maxArg)
    If Not hasMax Then maxArg = TQ_DEFAULT_MAX_BYTES
    If maxArg < TQ_MIN_BYTES Then TqRaiseParam "max_bytes must be an integer >= 4096"
    hasFmt = TqArgStr(args, "format", fmtArg)
    If hasFmt Then
        If fmtArg <> "json" And fmtArg <> "jsonl" Then TqRaiseParam "format must be ""json"" or ""jsonl"" (no other format exists)"
        fmtJsonl = (fmtArg = "jsonl")
    End If
    hasExpect = TqArgStr(args, "expect_head_sha256", expectArg)
    strictVal = JsonLib.JsonGetBool(args, "strict", False)
    countOnly = JsonLib.JsonGetBool(args, "count_only", False)
    Set fieldFilter = TqArgFields(args, "include_fields")
    ' ---------- 2. the filters echo (the A form's keys and value kinds) ----------
    Set filters = CreateObject("Scripting.Dictionary")
    If hasId Then
        filters("id") = idArg
    Else
        filters("id") = Null
    End If
    effectivePrefix = Null
    If hasPrefix Then effectivePrefix = prefixArg
    If hasProc Then effectivePrefix = procArg & "."
    filters("idPrefix") = effectivePrefix
    If hasProc Then
        filters("proc") = procArg
    Else
        filters("proc") = Null
    End If
    If hasPid Then
        filters("pid") = pidArg
    Else
        filters("pid") = Null
    End If
    Set fl = TqFieldList(fieldFilter)
    Set filters("includeFields") = fl
    ' ---------- 3. sink resolution (PURE: no filesystem access; D2/A.1.2) ----------
    sinkPath = Null
    pathFrom = Null
    If hasPath Then
        If Len(Trim$(pathArg)) > 0 Then
            sinkPath = pathArg
            pathFrom = "param"
        End If
    End If
    If IsNull(sinkPath) Then
        envT = Environ$("MCP_TRACE_FILE")
        If Len(envT) > 0 Then
            sinkPath = envT
            pathFrom = "env:MCP_TRACE_FILE"
        Else
            tempD = Environ$("TEMP")
            If Len(tempD) = 0 Then tempD = Environ$("TMP")
            If Len(tempD) > 0 Then
                If Right$(tempD, 1) = Chr$(92) Or Right$(tempD, 1) = "/" Then
                    sinkPath = tempD & "mcp_trace.jsonl"
                Else
                    sinkPath = tempD & Chr$(92) & "mcp_trace.jsonl"
                End If
                pathFrom = "env:TEMP"
            End If
        End If
    End If
    If hasSince Then
        Set receipt = TqBaseReceipt(sinkPath, pathFrom, sinceArg, filters, strictVal)
    Else
        Set receipt = TqBaseReceipt(sinkPath, pathFrom, Null, filters, strictVal)
    End If
    ' ---------- 4. the A form's ordered contract checks ----------
    If maxArg > TQ_MAX_BYTES Then
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("maxBytes") = maxArg
        failExtra("limitBytes") = TQ_MAX_BYTES
        failExtra("bytes") = Null
        failExtra("advice") = "max_bytes is a hard cap (ruling O6): request at most " & CStr(TQ_MAX_BYTES) & " bytes; the value is neither clamped to the default nor to the cap"
        Set TqQuery = TqFailFinished(receipt, TQ_R_TOO_LARGE, failExtra, started)
        Exit Function
    End If
    If hasSince Then
        If sinceArg < 0 Then
            Set cg = CreateObject("Scripting.Dictionary")
            cg("sinceOffset") = sinceArg
            Set failExtra = CreateObject("Scripting.Dictionary")
            Set failExtra("cursorGiven") = cg
            failExtra("fileBytes") = Null
            failExtra("alignedTo") = Null
            failExtra("headSha256") = Null
            failExtra("advice") = "since_offset must be >= 0 (byte offsets only; ruling A.1.6)"
            Set TqQuery = TqFailFinished(receipt, TQ_R_BAD_CURSOR, failExtra, started)
            Exit Function
        End If
    End If
    If IsNull(sinkPath) Then
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("advice") = "no sink path resolved: pass path, or set MCP_TRACE_FILE / TEMP for the process that runs this IDE. Nothing was created (design S6)"
        Set TqQuery = TqFailFinished(receipt, TQ_R_FILE_NOT_FOUND, failExtra, started)
        Exit Function
    End If
    sinkStr = CStr(sinkPath)
    policy = TqPathPolicy(sinkStr)
    If Len(policy) > 0 Then
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("policy") = policy
        failExtra("advice") = "nothing was opened: the path policy is checked before the first filesystem open (design S3)"
        Set TqQuery = TqFailFinished(receipt, TQ_R_REFUSED_PATH, failExtra, started)
        Exit Function
    End If
    ' ---------- 5. open + stat.  CreateFileW(OPEN_EXISTING) never creates a file. ----------
    If Not TqOpenRead(sinkStr, hFile, openErr) Then
        If openErr = TQ_ERR_FILE_NOT_FOUND Or openErr = TQ_ERR_PATH_NOT_FOUND Then
            Set fi = receipt("fileInfo")
            fi("exists") = False
            Set failExtra = CreateObject("Scripting.Dictionary")
            failExtra("advice") = "the sink file does not exist; run the traced program once, or check MCP_TRACE_FILE. Nothing was created (design S6)"
            Set TqQuery = TqFailFinished(receipt, TQ_R_FILE_NOT_FOUND, failExtra, started)
        Else
            Set failExtra = CreateObject("Scripting.Dictionary")
            Set failExtra("ioError") = TqIoError("win32", openErr, "CreateFileW failed; the path may be locked, denied or a directory")
            failExtra("advice") = "open failed; the file may be locked or inaccessible"
            Set TqQuery = TqFailFinished(receipt, TQ_R_IO_ERROR, failExtra, started)
        End If
        Exit Function
    End If
    Set fi = receipt("fileInfo")
    sizeD = TqFileSize(hFile)
    mtimeV = TqMtimeUtc(hFile)
    fi("exists") = True
    fi("size") = sizeD
    fi("mtime") = mtimeV
    If sizeD > TQ_MAX_FILE_BYTES Then
        CloseHandle hFile
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("bytes") = sizeD
        failExtra("limitBytes") = CLng(TQ_MAX_FILE_BYTES)
        failExtra("maxBytes") = Null
        failExtra("advice") = "the file exceeds the " & CStr(CLng(TQ_MAX_FILE_BYTES)) & " byte hard cap; it was NOT read at all (design S4)"
        Set TqQuery = TqFailFinished(receipt, TQ_R_TOO_LARGE, failExtra, started)
        Exit Function
    End If
    ' head hash: min(65536, size) bytes, lowercase hex (the A form's lowercase)
    wantLen = TQ_HEAD_HASH_BYTES
    If sizeD < CDbl(TQ_HEAD_HASH_BYTES) Then wantLen = CLng(sizeD)
    readErr = False
    headLen = TqReadAt(hFile, 0&, wantLen, headBuf, readErr)
    If readErr Then
        CloseHandle hFile
        Set failExtra = CreateObject("Scripting.Dictionary")
        Set failExtra("ioError") = TqIoError("win32", 0&, "the first 64 KiB could not be read")
        failExtra("advice") = "read failed; the file may have been locked or truncated concurrently"
        Set TqQuery = TqFailFinished(receipt, TQ_R_IO_ERROR, failExtra, started)
        Exit Function
    End If
    headSha = LCase$(Sha256.Sha256HexBytes(headBuf))
    fi("sha256Head64KiB") = headSha
    fi("headBytesRead") = headLen
    bomLen = TqBomLength(headBuf, headLen)
    If bomLen > 0 Then
        CloseHandle hFile
        Set probe = CreateObject("Scripting.Dictionary")
        probe("headHex") = TqHexLower(headBuf, 0&, headLen, 8&)
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("bomBytes") = bomLen
        Set failExtra("encodingProbe") = probe
        failExtra("advice") = "the file starts with a BOM; the sink never writes one - fail-closed instead of guessing the encoding (ruling O4)"
        Set TqQuery = TqFailFinished(receipt, TQ_R_ENCODING, failExtra, started)
        Exit Function
    End If
    ' expect_head_sha256: 64 hex characters, compared case-insensitively
    If hasExpect Then
        expectStr = LCase$(Trim$(expectArg))
        If Not TqIsHex64(expectStr) Then
            CloseHandle hFile
            Set cg = CreateObject("Scripting.Dictionary")
            cg("expectHeadSha256") = expectArg
            Set failExtra = CreateObject("Scripting.Dictionary")
            Set failExtra("cursorGiven") = cg
            failExtra("fileBytes") = sizeD
            failExtra("alignedTo") = Null
            failExtra("headSha256") = headSha
            failExtra("expectedHeadSha256") = Null
            failExtra("advice") = "expect_head_sha256 must be 64 hexadecimal characters"
            Set TqQuery = TqFailFinished(receipt, TQ_R_BAD_CURSOR, failExtra, started)
            Exit Function
        End If
        If expectStr <> headSha Then
            CloseHandle hFile
            Set cg = CreateObject("Scripting.Dictionary")
            cg("expectHeadSha256") = expectStr
            Set failExtra = CreateObject("Scripting.Dictionary")
            Set failExtra("cursorGiven") = cg
            failExtra("fileBytes") = sizeD
            failExtra("alignedTo") = Null
            failExtra("headSha256") = headSha
            failExtra("expectedHeadSha256") = expectStr
            failExtra("advice") = "the file identity guard failed: this is not the file the cursor was produced from (design 3.4: bad_cursor)"
            Set TqQuery = TqFailFinished(receipt, TQ_R_BAD_CURSOR, failExtra, started)
            Exit Function
        End If
    End If
    ' content-shape probe: refuse to report 0 hits for an arbitrary file
    Set probe = TqProbeShape(hFile, headBuf, headLen, sizeD)
    If Not CBool(probe("found")) Then
        CloseHandle hFile
        Set ffe = CreateObject("Scripting.Dictionary")
        ffe("linesChecked") = probe("linesChecked")
        ffe("traceShapedLines") = 0&
        Set failExtra = CreateObject("Scripting.Dictionary")
        failExtra("firstLineOffset") = probe("firstLineOffset")
        failExtra("excerpt") = probe("excerpt")
        Set failExtra("fileFormatEvidence") = ffe
        failExtra("advice") = "none of the first " & CStr(TQ_TRACE_PROBE_LINES) & " non-empty lines is a JSON object with a string id key; refusing to report 0 hits for an arbitrary file (design ACC-TQ11)"
        Set TqQuery = TqFailFinished(receipt, TQ_R_NOT_TRACE, failExtra, started)
        Exit Function
    End If
    ' ---------- 6. cursor placement (byte boundary is the ONLY cursor) ----------
    effectiveStart = 0&
    shrank = False
    If hasSince Then
        If CDbl(sinceArg) > sizeD Then
            shrank = True
        ElseIf sinceArg > 0 Then
            readErr = False
            prevLen = TqReadAt(hFile, sinceArg - 1, 1&, prevBuf, readErr)
            If readErr Then
                CloseHandle hFile
                Set failExtra = CreateObject("Scripting.Dictionary")
                Set failExtra("ioError") = TqIoError("win32", 0&, "the byte before since_offset could not be read")
                failExtra("advice") = "read failed; the file may have been locked or truncated concurrently"
                Set TqQuery = TqFailFinished(receipt, TQ_R_IO_ERROR, failExtra, started)
                Exit Function
            End If
            If (prevLen <> 1) Or (prevBuf(0) <> 10) Then
                alignedTo = TqNextLineOffset(hFile, sinceArg, sizeD)
                CloseHandle hFile
                Set cg = CreateObject("Scripting.Dictionary")
                cg("sinceOffset") = sinceArg
                Set failExtra = CreateObject("Scripting.Dictionary")
                Set failExtra("cursorGiven") = cg
                failExtra("fileBytes") = sizeD
                failExtra("alignedTo") = alignedTo
                failExtra("headSha256") = headSha
                failExtra("advice") = "since_offset is inside a line; ruling A.1.4 requires a line boundary (0 or right after LF) and makes a mid-line offset bad_cursor instead of re-aligning"
                Set TqQuery = TqFailFinished(receipt, TQ_R_BAD_CURSOR, failExtra, started)
                Exit Function
            End If
            effectiveStart = sinceArg
        End If
    End If
    ' ---------- 7. read the window ----------
    remainD = sizeD - CDbl(effectiveStart)
    If remainD < 0 Then remainD = 0
    windowD = remainD
    If CDbl(maxArg) < windowD Then windowD = CDbl(maxArg)
    windowLen = CLng(windowD)
    readErr = False
    bufLen = TqReadAt(hFile, effectiveStart, windowLen, buf, readErr)
    CloseHandle hFile
    If readErr Then
        Set failExtra = CreateObject("Scripting.Dictionary")
        Set failExtra("ioError") = TqIoError("win32", 0&, "the scan window could not be read")
        failExtra("advice") = "read failed; the file may have been locked or truncated concurrently"
        Set TqQuery = TqFailFinished(receipt, TQ_R_IO_ERROR, failExtra, started)
        Exit Function
    End If
    readHitEof = (bufLen < windowLen) Or (CDbl(effectiveStart) + CDbl(bufLen) >= sizeD)
    firstMatchOffset = -1
    Set scan = TqScanWindow(buf, bufLen, effectiveStart, sizeD, filters, limitArg, countOnly, strictVal, fieldFilter, fmtJsonl, readHitEof, firstMatchOffset)
    ' ---------- 8. strict failure: fail at the FIRST bad line (never consume it) ----------
    If Not IsNull(scan("strictReason")) Then
        consumed = CLng(scan("consumed"))
        endOffset = effectiveStart + consumed
        Set r = receipt("range")
        r("startOffset") = effectiveStart
        r("endOffset") = endOffset
        r("bytesRead") = endOffset - effectiveStart
        r("sha256") = TqRangeSha(buf, consumed)
        Set cn = receipt("counts")
        cn("linesScanned") = CLng(scan("scanned"))
        cn("eventsMatched") = CLng(scan("matched"))
        cn("eventsSkipped") = CLng(scan("skipped"))
        cn("badLines") = 1&
        cn("hits") = 0&
        Set bads = scan("badLines")
        Set sb = scan("strictBad")
        Set failExtra = CreateObject("Scripting.Dictionary")
        Set failExtra("badLines") = bads
        failExtra("badLineCount") = 1&
        Set failExtra("firstBadLine") = sb
        failExtra("advice") = "strict:true stops at the first bad line (ruling A.1.8); re-run without strict to skip and count it"
        If CStr(scan("strictReason")) = TQ_R_ENCODING Then
            Set eps = CreateObject("Scripting.Dictionary")
            eps("line") = sb("seq")
            eps("lineOffset") = sb("lineOffset")
            eps("sampleHex") = sb("sampleHex")
            failExtra("bomBytes") = 0&
            Set failExtra("encodingProbe") = eps
        End If
        Set TqQuery = TqFailFinished(receipt, CStr(scan("strictReason")), failExtra, started)
        Exit Function
    End If
    ' ---------- 9. success receipt ----------
    consumed = CLng(scan("consumed"))
    endOffset = effectiveStart + consumed
    Set r = receipt("range")
    r("startOffset") = effectiveStart
    r("endOffset") = endOffset
    r("bytesRead") = endOffset - effectiveStart
    r("sha256") = TqRangeSha(buf, consumed)
    truncatedBy = scan("truncatedBy")
    partialBytes = CLng(scan("partial"))
    reachedEof = CBool(scan("reachedEof"))
    Set cu = receipt("cursor")
    If hasSince Then
        cu("sinceOffset") = sinceArg
    Else
        cu("sinceOffset") = Null
    End If
    cu("nextOffset") = CLng(scan("nextOffset"))
    cu("partialLineBytes") = partialBytes
    cu("reachedEof") = reachedEof
    Set cn = receipt("counts")
    cn("linesScanned") = CLng(scan("scanned"))
    cn("eventsMatched") = CLng(scan("matched"))
    cn("eventsSkipped") = CLng(scan("skipped"))
    cn("badLines") = CLng(scan("badCount"))
    Set evs = scan("events")
    If countOnly Then
        cn("hits") = CLng(scan("matched"))
    Else
        cn("hits") = CLng(evs.Count)
    End If
    Set receipt("events") = evs
    Set bads = scan("badLines")
    Set receipt("badLines") = bads
    If IsNull(truncatedBy) Then
        receipt("truncated") = False
        receipt("truncatedBy") = Null
    Else
        receipt("truncated") = True
        receipt("truncatedBy") = truncatedBy
    End If
    Set warnings = New Collection
    If shrank Then warnings.Add TqWarning(TQ_W_FILE_SHRANK, Null, "since_offset " & CStr(sinceArg) & " is beyond the file size " & CStr(CLng(sizeD)) & "; the cursor was reset to 0 (design 4.3: a truncated/rotated file is a reset, not bad_cursor)")
    If CLng(scan("badCount")) > 0 Then warnings.Add TqWarning(TQ_W_BAD_SKIPPED, Null, CStr(CLng(scan("badCount"))) & " bad line(s) skipped (strict:false); badLines[] records the first " & CStr(bads.Count))
    If (partialBytes > 0) And reachedEof Then warnings.Add TqWarning(TQ_W_PARTIAL, Null, "the file ends with " & CStr(partialBytes) & " byte(s) that are not a complete line; they were not parsed and nextOffset stayed at their first byte")
    If CLng(scan("dupCount")) > 0 Then warnings.Add TqWarning(TQ_W_DUP_KEYS, CLng(scan("dupFirstSeq")), TqDupDetail(scan))
    If Not fieldFilter Is Nothing Then
        Set fsSeen = scan("fieldsSeen")
        detail = TqMissingFields(fieldFilter, fsSeen)
        If Len(detail) > 0 Then warnings.Add TqWarning(TQ_W_FIELDS_MISSING, Null, detail)
    End If
    If Not hasExpect Then warnings.Add TqWarning(TQ_W_HEAD_ABSENT, Null, "no expect_head_sha256 was supplied; fileInfo.sha256Head64KiB is returned so a caller can pass it back as a guard when the file identity matters")
    Set receipt("warnings") = warnings
    receipt("ok") = True
    Set TqQuery = TqFinish(receipt, started)
End Function

' ============================================================================
' small receipt helpers used by TqQuery
' ============================================================================
Private Function TqFailFinished(ByVal receipt As Object, ByVal reason As String, ByVal extras As Object, ByVal started As Long) As Object
    Dim r As Object
    Set r = TqFail(receipt, reason, extras)
    Set TqFailFinished = TqFinish(r, started)
End Function

Private Function TqWarning(ByVal code As String, ByVal lineV As Variant, ByVal detail As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("code") = code
    d("line") = lineV
    d("detail") = detail
    Set TqWarning = d
End Function

' io_error evidence: the Win32 number and a plain description, never swallowed.
Private Function TqIoError(ByVal kind As String, ByVal win32Code As Long, ByVal msg As String) As Object
    Dim d As Object
    Dim ve As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("type") = kind
    If win32Code <> 0 Then d("win32Error") = win32Code
    Set ve = CreateObject("Scripting.Dictionary")
    ve("number") = win32Code
    ve("description") = msg
    Set d("vbError") = ve
    d("message") = msg
    Set TqIoError = d
End Function

Private Function TqBadLine(ByVal seqNo As Long, ByVal lineOffset As Long, ByVal lineBytes As Long, ByVal errKind As String, ByVal excerpt As String, ByVal sampleHex As String) As Object
    Dim d As Object
    Set d = CreateObject("Scripting.Dictionary")
    d("seq") = seqNo
    d("lineOffset") = lineOffset
    d("lineBytes") = lineBytes
    d("error") = errKind
    d("excerpt") = excerpt
    d("sampleHex") = sampleHex
    Set TqBadLine = d
End Function

' duplicate_keys detail, built from the scan's dupRecords (A form wording).
Private Function TqDupDetail(ByVal scan As Object) As String
    Dim recs As Object
    Dim allKeys As Object
    Dim rec As Object
    Dim keys As Object
    Dim i As Long
    Dim k As Variant
    Dim seqs As String
    Dim keyList As String
    Dim c As Long
    Set recs = scan("dupRecords")
    Set allKeys = CreateObject("Scripting.Dictionary")
    seqs = ""
    For i = 1 To recs.Count
        Set rec = recs(i)
        If i <= TQ_MAX_DUP_LINES Then
            If Len(seqs) > 0 Then seqs = seqs & ","
            seqs = seqs & CStr(rec("seq"))
        End If
        Set keys = rec("keys")
        For Each k In keys
            If Not allKeys.Exists(CStr(k)) Then allKeys(CStr(k)) = True
        Next k
    Next i
    c = 0
    keyList = ""
    For Each k In allKeys.Keys
        If c < 20 Then
            If Len(keyList) > 0 Then keyList = keyList & ","
            keyList = keyList & CStr(k)
            c = c + 1
        End If
    Next k
    TqDupDetail = CStr(recs.Count) & " event line(s) repeat a JSON key (last value kept); first line numbers: " & seqs & "; keys: " & keyList
End Function

' fields_missing detail: named filters that never occurred in a matched event.
Private Function TqMissingFields(ByVal fieldDict As Object, ByVal fieldsSeen As Object) As String
    Dim k As Variant
    Dim missing As String
    missing = ""
    For Each k In fieldDict.Keys
        If CStr(k) <> "id" And CStr(k) <> "pid" Then
            If Not fieldsSeen.Exists(CStr(k)) Then
                If Len(missing) > 0 Then missing = missing & ","
                missing = missing & CStr(k)
            End If
        End If
    Next k
    If Len(missing) = 0 Then
        TqMissingFields = ""
    Else
        TqMissingFields = "include_fields name(s) not present in any matched event: " & missing
    End If
End Function

' ============================================================================
' path policy (design S3 / ruling O9 / ruling O11).  Runs BEFORE any open:
' only shape checks plus a GetAttr probe, which never opens the file content.
' ============================================================================
Private Function TqPathPolicy(ByVal path As String) As String
    Dim attr As Long
    Dim errN As Long
    Dim bs As String
    TqPathPolicy = ""
    bs = Chr$(92)
    If Left$(path, 2) = bs & bs Or Left$(path, 2) = "//" Then
        TqPathPolicy = "network or device-namespace path refused: path must be a LOCAL absolute path (UNC and // paths are refused_path)"
        Exit Function
    End If
    If Not TqIsAbsolute(path) Then
        TqPathPolicy = "relative path refused: path must be an absolute local path"
        Exit Function
    End If
    On Error Resume Next
    Err.Clear
    attr = 0&
    attr = GetAttr(path)
    errN = Err.Number
    Err.Clear
    On Error GoTo 0
    If errN = 0 Then
        If (attr And vbDirectory) <> 0 Then
            TqPathPolicy = "refused: the path exists but is not a regular file (directory or device)"
        End If
    End If
End Function

' Mirrors the A form's os.path.isabs on Windows: drive-rooted (X:) needs the
' separator at position 3; a single leading separator is rooted on the current
' drive (two leading separators were already refused as UNC).  A drive-relative
' shape such as "C:foo" is NOT absolute and is refused.
Private Function TqIsAbsolute(ByVal path As String) As Boolean
    Dim bs As String
    TqIsAbsolute = False
    bs = Chr$(92)
    If Len(path) >= 2 Then
        If Mid$(path, 2, 1) = ":" Then
            If Len(path) >= 3 Then
                If Mid$(path, 3, 1) = bs Or Mid$(path, 3, 1) = "/" Then TqIsAbsolute = True
            End If
            Exit Function
        End If
    End If
    If Left$(path, 1) = bs Or Left$(path, 1) = "/" Then TqIsAbsolute = True
End Function

' ============================================================================
' file primitives.  CreateFileW(OPEN_EXISTING) never creates the target, which
' is the whole reason VB's "Open ... For Binary Access Read" is not used here.
' ============================================================================
Private Function TqOpenRead(ByVal path As String, ByRef hFile As Long, ByRef winErr As Long) As Boolean
    Dim h As Long
    TqOpenRead = False
    hFile = TQ_INVALID_HANDLE
    winErr = 0
    On Error Resume Next
    Err.Clear
    h = CreateFileW(StrPtr(path), TQ_FILE_READ_DATA Or TQ_FILE_READ_ATTRIBUTES, TQ_SHARE_ALL, 0&, TQ_OPEN_EXISTING, 0&, 0&)
    winErr = Err.LastDllError
    On Error GoTo 0
    If h = TQ_INVALID_HANDLE Then Exit Function
    hFile = h
    winErr = 0
    TqOpenRead = True
End Function

' File size as a Double, from the open handle (GetFileSize returns a DWORD in
' a Long, so high values arrive negative and are corrected here).
Private Function TqFileSize(ByVal hFile As Long) As Double
    Dim lo As Long
    Dim hi As Long
    Dim lowD As Double
    Dim highD As Double
    hi = 0&
    lo = GetFileSize(hFile, hi)
    lowD = CDbl(lo)
    If lo < 0 Then lowD = lowD + 4294967296#
    highD = CDbl(hi)
    If hi < 0 Then highD = highD + 4294967296#
    TqFileSize = lowD + highD * 4294967296#
End Function

' REAL UTC last-write time (ruling BATCH-O4), or Null when it cannot be read -
' never a local-time guess and never a fabricated value.
Private Function TqMtimeUtc(ByVal hFile As Long) As Variant
    Dim ftC As TqFileTime
    Dim ftA As TqFileTime
    Dim ftW As TqFileTime
    Dim st As TqSystemTime
    TqMtimeUtc = Null
    If GetFileTime(hFile, ftC, ftA, ftW) = 0 Then Exit Function
    If FileTimeToSystemTime(ftW, st) = 0 Then Exit Function
    TqMtimeUtc = Format$(st.wYear, "0000") & "-" & Format$(st.wMonth, "00") & "-" & Format$(st.wDay, "00") & "T" & Format$(st.wHour, "00") & ":" & Format$(st.wMinute, "00") & ":" & Format$(st.wSecond, "00") & "Z"
End Function

' Read up to want bytes at offset.  Returns the bytes actually read; a SHORT
' read is legal (the file may have been truncated concurrently) and the caller
' decides what that means.  errFlag is set only for a real I/O error.
Private Function TqReadAt(ByVal hFile As Long, ByVal offset As Long, ByVal want As Long, ByRef buf() As Byte, ByRef errFlag As Boolean) As Long
    Dim high As Long
    Dim got As Long
    Dim total As Long
    Dim rc As Long
    Dim lastErr As Long
    errFlag = False
    TqReadAt = 0
    Erase buf
    If want <= 0 Then Exit Function
    high = 0&
    On Error Resume Next
    Err.Clear
    rc = SetFilePointer(hFile, offset, high, TQ_FILE_BEGIN)
    lastErr = Err.LastDllError
    On Error GoTo 0
    If rc = -1& Then
        If lastErr <> 0 Then
            errFlag = True
            Exit Function
        End If
    End If
    ReDim buf(0 To want - 1)
    total = 0&
    Do While total < want
        got = 0&
        On Error Resume Next
        Err.Clear
        rc = ReadFile(hFile, buf(total), want - total, got, 0&)
        lastErr = Err.LastDllError
        On Error GoTo 0
        If rc = 0 Then
            If lastErr = 0 Or lastErr = TQ_ERR_HANDLE_EOF Then Exit Do
            errFlag = True
            Exit Do
        End If
        If got <= 0 Then Exit Do
        total = total + got
    Loop
    If errFlag Then
        TqReadAt = total
        Exit Function
    End If
    If total = 0 Then
        Erase buf
    ElseIf total < want Then
        ReDim Preserve buf(0 To total - 1)
    End If
    TqReadAt = total
End Function

Private Function TqByteCount(ByRef b() As Byte) As Long
    On Error Resume Next
    TqByteCount = 0
    TqByteCount = UBound(b) - LBound(b) + 1
    On Error GoTo 0
End Function

' The 5 BOM prefixes the A form refuses, LONGEST FIRST (UTF-32LE starts with
' the UTF-16LE prefix, so order matters).
Private Function TqBomLength(ByRef b() As Byte, ByVal n As Long) As Long
    TqBomLength = 0
    If n >= 3 Then
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then
            TqBomLength = 3
            Exit Function
        End If
    End If
    If n >= 4 Then
        If b(0) = &HFF And b(1) = &HFE And b(2) = 0 And b(3) = 0 Then
            TqBomLength = 4
            Exit Function
        End If
        If b(0) = 0 And b(1) = 0 And b(2) = &HFE And b(3) = &HFF Then
            TqBomLength = 4
            Exit Function
        End If
    End If
    If n >= 2 Then
        If b(0) = &HFF And b(1) = &HFE Then
            TqBomLength = 2
            Exit Function
        End If
        If b(0) = &HFE And b(1) = &HFF Then
            TqBomLength = 2
            Exit Function
        End If
    End If
End Function

Private Function TqIsHex64(ByVal s As String) As Boolean
    Dim i As Long
    TqIsHex64 = False
    If Len(s) <> 64 Then Exit Function
    For i = 1 To 64
        If InStr("0123456789abcdef", Mid$(s, i, 1)) = 0 Then Exit Function
    Next i
    TqIsHex64 = True
End Function

Private Function TqHexLower(ByRef b() As Byte, ByVal startPos As Long, ByVal totalLen As Long, ByVal maxLen As Long) As String
    Dim i As Long
    Dim n As Long
    Dim s As String
    Dim h As String
    n = totalLen
    If n > maxLen Then n = maxLen
    s = ""
    For i = 0 To n - 1
        h = Hex$(b(startPos + i))
        If Len(h) < 2 Then h = "0" & h
        s = s & LCase$(h)
    Next i
    TqHexLower = s
End Function

' range.sha256: the digest of EXACTLY the consumed bytes (an empty window hashes
' the empty byte sequence, exactly like the A form).
Private Function TqRangeSha(ByRef buf() As Byte, ByVal consumed As Long) As String
    Dim slice() As Byte
    Dim n As Long
    n = TqByteCount(buf)
    If consumed <= 0 Then
        TqRangeSha = LCase$(Sha256.Sha256HexBytes(slice))
        Exit Function
    End If
    If consumed >= n Then
        TqRangeSha = LCase$(Sha256.Sha256HexBytes(buf))
        Exit Function
    End If
    ReDim slice(0 To consumed - 1)
    CopyMemory slice(0), buf(0), consumed
    TqRangeSha = LCase$(Sha256.Sha256HexBytes(slice))
End Function

' ============================================================================
' strict UTF-8 decoding of ONE line body (ruling O4: no replacement, no ANSI
' fallback).  False = the line is not valid UTF-8 and is a bad line.
' ============================================================================
Private Function TqDecodeBody(ByRef src() As Byte, ByVal startPos As Long, ByVal bodyLen As Long, ByRef outText As String) As Boolean
    Dim tmp() As Byte
    outText = ""
    TqDecodeBody = True
    If bodyLen <= 0 Then Exit Function
    ReDim tmp(0 To bodyLen - 1)
    CopyMemory tmp(0), src(startPos), bodyLen
    If Utf8.TryUtf8Decode(tmp, outText) Then
        TqDecodeBody = True
    Else
        outText = ""
        TqDecodeBody = False
    End If
End Function

' Display-only excerpt: UTF-8 text when it decoded, else a lossless byte-per-
' character rendering (the A form's latin-1 fallback), capped at 200 characters.
Private Function TqExcerpt(ByVal decoded As Boolean, ByVal decodedText As String, ByRef src() As Byte, ByVal startPos As Long, ByVal bodyLen As Long) As String
    Dim n As Long
    Dim tmp() As Byte
    If decoded Then
        If Len(decodedText) > TQ_EXCERPT_MAX Then
            TqExcerpt = Left$(decodedText, TQ_EXCERPT_MAX)
        Else
            TqExcerpt = decodedText
        End If
        Exit Function
    End If
    n = bodyLen
    If n > TQ_EXCERPT_MAX Then n = TQ_EXCERPT_MAX
    If n <= 0 Then
        TqExcerpt = ""
        Exit Function
    End If
    ReDim tmp(0 To n - 1)
    CopyMemory tmp(0), src(startPos), n
    TqExcerpt = Utf8.BytesToStr(tmp)
End Function

' ============================================================================
' JSON lexical validation.  JsonLib.JsonParse is the object builder, but it is
' permissive in places where the A form (Pythons json.loads) is strict:
'   * it accepts a bare TAB / other control character inside a string literal;
'   * its number scanner uses Val(), so "1.2.3" and "+1" would be accepted;
'   * it ignores trailing garbage after the first value;
'   * it accepts any four characters after a u escape.
' So every line is FIRST validated by the scanner below (JSON grammar, strings
' escaped, numbers strict, one value and only whitespace after it, at most 64
' nesting levels) and only then handed to JsonLib.  The scanner also reports the
' top-level key facts the tool needs: is this an object, is id a string, is pid
' an INTEGER literal (not 1.0 / 1e3), and which TOP-LEVEL keys repeat (D7).
' ============================================================================
Private Sub TqSkipWs(ByVal s As String, ByRef pos As Long)
    Dim n As Long
    Dim c As Long
    n = Len(s)
    Do While pos <= n
        c = AscW(Mid$(s, pos, 1)) And &HFFFF&
        If c = 32 Or c = 9 Or c = 10 Or c = 13 Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
End Sub

' pos points at the opening quote; on success it points just past the closing
' quote.  Rejects raw control characters inside the literal and bad escapes.
Private Function TqSkipString(ByVal s As String, ByRef pos As Long) As Boolean
    Dim n As Long
    Dim c As String
    Dim esc As String
    Dim cc As Long
    Dim k As Long
    TqSkipString = False
    n = Len(s)
    If pos > n Then Exit Function
    If Mid$(s, pos, 1) <> Chr$(34) Then Exit Function
    pos = pos + 1
    Do While pos <= n
        c = Mid$(s, pos, 1)
        cc = AscW(c) And &HFFFF&
        If cc = 34 Then
            pos = pos + 1
            TqSkipString = True
            Exit Function
        End If
        If cc = 92 Then
            pos = pos + 1
            If pos > n Then Exit Function
            esc = Mid$(s, pos, 1)
            If esc = Chr$(34) Or esc = Chr$(92) Or esc = "/" Then
                pos = pos + 1
            ElseIf esc = "b" Or esc = "f" Or esc = "n" Or esc = "r" Or esc = "t" Then
                pos = pos + 1
            ElseIf esc = "u" Then
                If pos + 4 > n Then Exit Function
                For k = 1 To 4
                    If InStr("0123456789abcdefABCDEF", Mid$(s, pos + k, 1)) = 0 Then Exit Function
                Next k
                pos = pos + 5
            Else
                Exit Function
            End If
        Else
            If cc < 32 Then Exit Function
            pos = pos + 1
        End If
    Loop
End Function

' Strict JSON number grammar (a leading +, a leading zero, a bare dot or a
' missing exponent digit are all rejected).  kind is TQ_K_INT for an integer
' literal, TQ_K_NUM when a fraction or an exponent is present.
Private Function TqSkipNumber(ByVal s As String, ByRef pos As Long, ByRef kind As Long) As Boolean
    Dim n As Long
    Dim c As String
    Dim isInt As Boolean
    TqSkipNumber = False
    kind = TQ_K_INT
    n = Len(s)
    isInt = True
    If pos <= n Then
        If Mid$(s, pos, 1) = "-" Then pos = pos + 1
    End If
    If pos > n Then Exit Function
    c = Mid$(s, pos, 1)
    If c = "0" Then
        pos = pos + 1
        If pos <= n Then
            c = Mid$(s, pos, 1)
            If c >= "0" And c <= "9" Then Exit Function
        End If
    ElseIf c >= "1" And c <= "9" Then
        Do While pos <= n
            c = Mid$(s, pos, 1)
            If c < "0" Or c > "9" Then Exit Do
            pos = pos + 1
        Loop
    Else
        Exit Function
    End If
    If pos <= n Then
        If Mid$(s, pos, 1) = "." Then
            isInt = False
            pos = pos + 1
            If pos > n Then Exit Function
            c = Mid$(s, pos, 1)
            If c < "0" Or c > "9" Then Exit Function
            Do While pos <= n
                c = Mid$(s, pos, 1)
                If c < "0" Or c > "9" Then Exit Do
                pos = pos + 1
            Loop
        End If
    End If
    If pos <= n Then
        c = Mid$(s, pos, 1)
        If c = "e" Or c = "E" Then
            isInt = False
            pos = pos + 1
            If pos <= n Then
                c = Mid$(s, pos, 1)
                If c = "+" Or c = "-" Then pos = pos + 1
            End If
            If pos > n Then Exit Function
            c = Mid$(s, pos, 1)
            If c < "0" Or c > "9" Then Exit Function
            Do While pos <= n
                c = Mid$(s, pos, 1)
                If c < "0" Or c > "9" Then Exit Do
                pos = pos + 1
            Loop
        End If
    End If
    If isInt Then
        kind = TQ_K_INT
    Else
        kind = TQ_K_NUM
    End If
    TqSkipNumber = True
End Function

' Skip exactly one JSON value, validating it.  Recursion is bounded by
' TQ_MAX_JSON_DEPTH so a hostile line can never overflow the stack.
Private Function TqSkipValue(ByVal s As String, ByRef pos As Long, ByVal depth As Long, ByRef kind As Long) As Boolean
    Dim c As String
    Dim n As Long
    TqSkipValue = False
    kind = 0&
    If depth > TQ_MAX_JSON_DEPTH Then Exit Function
    n = Len(s)
    If n <= 0 Then Exit Function
    TqSkipWs s, pos
    If pos > n Then Exit Function
    c = Mid$(s, pos, 1)
    If c = "{" Then
        pos = pos + 1
        TqSkipWs s, pos
        If pos > n Then Exit Function
        If Mid$(s, pos, 1) = "}" Then
            pos = pos + 1
            kind = TQ_K_OBJECT
            TqSkipValue = True
            Exit Function
        End If
        Do
            TqSkipWs s, pos
            If Not TqSkipString(s, pos) Then Exit Function
            TqSkipWs s, pos
            If pos > n Then Exit Function
            If Mid$(s, pos, 1) <> ":" Then Exit Function
            pos = pos + 1
            If Not TqSkipValue(s, pos, depth + 1, kind) Then Exit Function
            TqSkipWs s, pos
            If pos > n Then Exit Function
            If Mid$(s, pos, 1) = "," Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        TqSkipWs s, pos
        If pos > n Then Exit Function
        If Mid$(s, pos, 1) <> "}" Then Exit Function
        pos = pos + 1
        kind = TQ_K_OBJECT
        TqSkipValue = True
    ElseIf c = "[" Then
        pos = pos + 1
        TqSkipWs s, pos
        If pos > n Then Exit Function
        If Mid$(s, pos, 1) = "]" Then
            pos = pos + 1
            kind = TQ_K_ARRAY
            TqSkipValue = True
            Exit Function
        End If
        Do
            If Not TqSkipValue(s, pos, depth + 1, kind) Then Exit Function
            TqSkipWs s, pos
            If pos > n Then Exit Function
            If Mid$(s, pos, 1) = "," Then
                pos = pos + 1
            Else
                Exit Do
            End If
        Loop
        TqSkipWs s, pos
        If pos > n Then Exit Function
        If Mid$(s, pos, 1) <> "]" Then Exit Function
        pos = pos + 1
        kind = TQ_K_ARRAY
        TqSkipValue = True
    ElseIf c = Chr$(34) Then
        If Not TqSkipString(s, pos) Then Exit Function
        kind = TQ_K_STRING
        TqSkipValue = True
    ElseIf c = "t" Then
        If Mid$(s, pos, 4) = "true" Then
            pos = pos + 4
            kind = TQ_K_LITERAL
            TqSkipValue = True
        End If
    ElseIf c = "f" Then
        If Mid$(s, pos, 5) = "false" Then
            pos = pos + 5
            kind = TQ_K_LITERAL
            TqSkipValue = True
        End If
    ElseIf c = "n" Then
        If Mid$(s, pos, 4) = "null" Then
            pos = pos + 4
            kind = TQ_K_LITERAL
            TqSkipValue = True
        End If
    Else
        If TqSkipNumber(s, pos, kind) Then TqSkipValue = True
    End If
End Function

' Decode one JSON string literal key.  startPos = the opening quote, endPos =
' ONE PAST the closing quote (the convention TqSkipString leaves behind), so
' the content characters are startPos+1 .. endPos-2.  Only used for keys.
Private Function TqDecodeKey(ByVal s As String, ByVal startPos As Long, ByVal endPos As Long) As String
    Dim i As Long
    Dim out As String
    Dim c As String
    Dim cc As Long
    Dim hex4 As String
    i = startPos + 1
    out = ""
    Do While i < endPos - 1
        c = Mid$(s, i, 1)
        cc = AscW(c) And &HFFFF&
        If cc = 92 Then
            i = i + 1
            If i >= endPos Then Exit Do
            c = Mid$(s, i, 1)
            If c = "b" Then
                out = out & Chr$(8)
            ElseIf c = "f" Then
                out = out & Chr$(12)
            ElseIf c = "n" Then
                out = out & vbLf
            ElseIf c = "r" Then
                out = out & Chr$(13)
            ElseIf c = "t" Then
                out = out & vbTab
            ElseIf c = "u" Then
                hex4 = Mid$(s, i + 1, 4)
                out = out & ChrW$(CLng("&H" & hex4))
                i = i + 4
            Else
                out = out & c
            End If
            i = i + 1
        Else
            out = out & c
            i = i + 1
        End If
    Loop
    TqDecodeKey = out
End Function

' Top-level object walk: fills idOk ('"id" holds a JSON string'), pidOk ('"pid"
' holds an INTEGER literal') and dupKeys (every top-level key that repeats).
Private Function TqScanTopObject(ByVal s As String, ByRef pos As Long, ByRef idOk As Boolean, ByRef pidOk As Boolean, ByRef dupKeys As Object) As Boolean
    Dim n As Long
    Dim keys As Object
    Dim keyText As String
    Dim keyStart As Long
    Dim vkind As Long
    Dim c As String
    TqScanTopObject = False
    Set keys = CreateObject("Scripting.Dictionary")
    idOk = False
    pidOk = False
    n = Len(s)
    pos = pos + 1
    TqSkipWs s, pos
    If pos > n Then Exit Function
    If Mid$(s, pos, 1) = "}" Then
        pos = pos + 1
        TqScanTopObject = True
        Exit Function
    End If
    Do
        TqSkipWs s, pos
        keyStart = pos
        If Not TqSkipString(s, pos) Then Exit Function
        keyText = TqDecodeKey(s, keyStart, pos)
        TqSkipWs s, pos
        If pos > n Then Exit Function
        If Mid$(s, pos, 1) <> ":" Then Exit Function
        pos = pos + 1
        TqSkipWs s, pos
        If Not TqSkipValue(s, pos, 2&, vkind) Then Exit Function
        If keys.Exists(keyText) Then
            dupKeys.Add keyText
        Else
            keys(keyText) = True
        End If
        If keyText = "id" Then
            idOk = (vkind = TQ_K_STRING)
        ElseIf keyText = "pid" Then
            pidOk = (vkind = TQ_K_INT)
        End If
        TqSkipWs s, pos
        If pos > n Then Exit Function
        c = Mid$(s, pos, 1)
        If c = "," Then
            pos = pos + 1
        Else
            Exit Do
        End If
    Loop
    TqSkipWs s, pos
    If pos > n Then Exit Function
    If Mid$(s, pos, 1) <> "}" Then Exit Function
    pos = pos + 1
    TqScanTopObject = True
End Function

' TRUE only when s is exactly ONE valid JSON value with nothing but whitespace
' after it.  startsObj / idOk / pidOk / dupKeys describe a top-level object.
Private Function TqJsonValidate(ByVal s As String, ByRef startsObj As Boolean, ByRef idOk As Boolean, ByRef pidOk As Boolean, ByRef dupKeys As Object) As Boolean
    Dim pos As Long
    Dim kind As Long
    Dim n As Long
    Set dupKeys = New Collection
    startsObj = False
    idOk = False
    pidOk = False
    TqJsonValidate = False
    n = Len(s)
    If n <= 0 Then Exit Function
    pos = 1&
    TqSkipWs s, pos
    If pos > n Then Exit Function
    If Mid$(s, pos, 1) = "{" Then
        startsObj = True
        If Not TqScanTopObject(s, pos, idOk, pidOk, dupKeys) Then Exit Function
    Else
        If Not TqSkipValue(s, pos, 1&, kind) Then Exit Function
    End If
    TqSkipWs s, pos
    If pos <= n Then Exit Function
    TqJsonValidate = True
End Function

' Build the Scripting.Dictionary of a validated object line.  JsonLib may raise
' (bad unicode escape and friends), so the call is wrapped: a raise is a bad line.
Private Function TqParseObject(ByVal s As String, ByRef outObj As Object) As Boolean
    Dim v As Variant
    Dim parsed As Boolean
    Dim eN As Long
    Set outObj = Nothing
    TqParseObject = False
    On Error Resume Next
    Err.Clear
    parsed = JsonLib.JsonParse(s, v)
    eN = Err.Number
    Err.Clear
    On Error GoTo 0
    If Not parsed Then Exit Function
    If eN <> 0 Then Exit Function
    If Not IsObject(v) Then Exit Function
    If TypeName(v) <> "Dictionary" Then Exit Function
    Set outObj = v
    TqParseObject = True
End Function

' A line is trace shaped when it is a JSON object with a STRING id key (the A
' form's _is_trace_shaped; pid is NOT required for the shape probe).
Private Function TqIsTraceShaped(ByVal text As String) As Boolean
    Dim startsObj As Boolean
    Dim idOk As Boolean
    Dim pidOk As Boolean
    Dim dupKeys As Object
    TqIsTraceShaped = False
    If Not TqJsonValidate(text, startsObj, idOk, pidOk, dupKeys) Then Exit Function
    If Not startsObj Then Exit Function
    If Not idOk Then Exit Function
    TqIsTraceShaped = True
End Function

' ============================================================================
' content-shape probe: none of the first 64 non-empty lines is trace shaped =>
' not_a_trace_file (never "0 hits" for an arbitrary file).  When the 64 KiB head
' is too small the probe re-reads up to 1 MiB, exactly like the A form.
' ============================================================================
Private Function TqProbeShape(ByVal hFile As Long, ByRef headBuf() As Byte, ByVal headLen As Long, ByVal fileSize As Double) As Object
    Dim res As Object
    Dim wantD As Double
    Dim probeBuf() As Byte
    Dim probeLen As Long
    Dim errFlag As Boolean
    Set res = TqProbeBuffer(headBuf, headLen)
    If CBool(res("found")) Then
        Set TqProbeShape = res
        Exit Function
    End If
    If CLng(res("linesChecked")) < TQ_TRACE_PROBE_LINES Then
        If fileSize > CDbl(headLen) Then
            wantD = fileSize
            If wantD > CDbl(TQ_TRACE_PROBE_MAX_BYTES) Then wantD = CDbl(TQ_TRACE_PROBE_MAX_BYTES)
            errFlag = False
            probeLen = TqReadAt(hFile, 0&, CLng(wantD), probeBuf, errFlag)
            If Not errFlag Then Set res = TqProbeBuffer(probeBuf, probeLen)
        End If
    End If
    Set TqProbeShape = res
End Function

Private Function TqProbeBuffer(ByRef buf() As Byte, ByVal bufLen As Long) As Object
    Dim res As Object
    Dim pos As Long
    Dim nlPos As Long
    Dim bufIdx As Long
    Dim bodyLen As Long
    Dim bodyText As String
    Dim decodeOk As Boolean
    Dim checked As Long
    Set res = CreateObject("Scripting.Dictionary")
    res("found") = False
    res("firstLineOffset") = Null
    res("excerpt") = Null
    res("linesChecked") = 0&
    pos = 0&
    checked = 0&
    Do While pos < bufLen
        nlPos = -1&
        bufIdx = pos
        Do While bufIdx < bufLen
            If buf(bufIdx) = 10 Then
                nlPos = bufIdx
                Exit Do
            End If
            bufIdx = bufIdx + 1
        Loop
        If nlPos < 0 Then
            bodyLen = bufLen - pos
        Else
            bodyLen = nlPos - pos
        End If
        If bodyLen > 0 Then
            If buf(pos + bodyLen - 1) = 13 Then bodyLen = bodyLen - 1
        End If
        decodeOk = TqDecodeBody(buf, pos, bodyLen, bodyText)
        If decodeOk Then
            If Len(Trim$(bodyText)) > 0 Then
                checked = checked + 1
                If IsNull(res("firstLineOffset")) Then
                    res("firstLineOffset") = pos
                    res("excerpt") = TqExcerpt(True, bodyText, buf, pos, bodyLen)
                End If
                If TqIsTraceShaped(bodyText) Then
                    res("found") = True
                    res("linesChecked") = checked
                    Set TqProbeBuffer = res
                    Exit Function
                End If
                If checked >= TQ_TRACE_PROBE_LINES Then Exit Do
            End If
        Else
            checked = checked + 1
            If IsNull(res("firstLineOffset")) Then
                res("firstLineOffset") = pos
                res("excerpt") = TqExcerpt(False, "", buf, pos, bodyLen)
            End If
            If checked >= TQ_TRACE_PROBE_LINES Then Exit Do
        End If
        If nlPos < 0 Then Exit Do
        pos = nlPos + 1
    Loop
    res("linesChecked") = checked
    Set TqProbeBuffer = res
End Function

' Evidence only: the first byte after the next LF at/after offset (or Null).
' Bounded by TQ_ALIGN_SCAN_MAX bytes so a pathological single line cannot turn a
' bad_cursor receipt into an unbounded read.
Private Function TqNextLineOffset(ByVal hFile As Long, ByVal offset As Long, ByVal fileSize As Double) As Variant
    Dim pos As Long
    Dim scanned As Long
    Dim chunk As Long
    Dim buf() As Byte
    Dim got As Long
    Dim i As Long
    Dim errFlag As Boolean
    TqNextLineOffset = Null
    pos = offset
    scanned = 0&
    Do While CDbl(pos) < fileSize
        If scanned >= TQ_ALIGN_SCAN_MAX Then Exit Function
        chunk = TQ_HEAD_HASH_BYTES
        If fileSize - CDbl(pos) < CDbl(chunk) Then chunk = CLng(fileSize - CDbl(pos))
        errFlag = False
        got = TqReadAt(hFile, pos, chunk, buf, errFlag)
        If errFlag Then Exit Function
        If got <= 0 Then Exit Function
        For i = 0 To got - 1
            If buf(i) = 10 Then
                TqNextLineOffset = pos + i + 1
                Exit Function
            End If
        Next i
        pos = pos + got
        scanned = scanned + got
    Loop
End Function

' ============================================================================
' the scan.  Consumes COMPLETE lines only; the (limit+1)-th matching line is
' never consumed, so nextOffset always points at the next byte to read.
' ============================================================================
Private Function TqScanWindow(ByRef buf() As Byte, ByVal bufLen As Long, ByVal baseOffset As Long, ByVal fileSize As Double, ByVal filters As Object, ByVal limitVal As Long, ByVal countOnly As Boolean, ByVal strictVal As Boolean, ByVal fieldFilter As Object, ByVal includeRaw As Boolean, ByVal readHitEof As Boolean, ByRef firstMatchOffset As Long) As Object
    Dim res As Object
    Dim events As Object
    Dim badLines As Object
    Dim dupRecs As Object
    Dim fieldsSeen As Object
    Dim scanned As Long
    Dim matched As Long
    Dim skipped As Long
    Dim badCount As Long
    Dim seq As Long
    Dim pos As Long
    Dim nlPos As Long
    Dim bufIdx As Long
    Dim lineStart As Long
    Dim lineBytes As Long
    Dim bodyLen As Long
    Dim bodyText As String
    Dim decodeOk As Boolean
    Dim validJson As Boolean
    Dim startsObj As Boolean
    Dim idOk As Boolean
    Dim pidOk As Boolean
    Dim dupKeys As Object
    Dim errKind As String
    Dim excerpt As String
    Dim sampleHex As String
    Dim obj As Object
    Dim ev As Object
    Dim fields As Object
    Dim isMatch As Boolean
    Dim strictHit As Boolean
    Dim strictReason As Variant
    Dim strictBad As Variant
    Dim sbo As Object
    Dim badEntry As Object
    Dim truncatedBy As Variant
    Dim consumed As Long
    Dim nextOff As Long
    Dim reachedEof As Boolean
    Dim partial As Long
    Dim kk As Variant
    Dim rec As Object
    Dim wId As Variant
    Dim wPrefix As Variant
    Dim wPid As Variant
    Dim idText As String
    Dim prefixText As String
    Dim pidD As Double
    Dim wantPid As Double
    Set events = New Collection
    Set badLines = New Collection
    Set dupRecs = New Collection
    Set fieldsSeen = CreateObject("Scripting.Dictionary")
    scanned = 0&
    matched = 0&
    skipped = 0&
    badCount = 0&
    seq = 0&
    pos = 0&
    strictHit = False
    strictReason = Null
    strictBad = Null
    truncatedBy = Null
    firstMatchOffset = -1&
    wId = filters("id")
    wPrefix = filters("idPrefix")
    wPid = filters("pid")
    Do While pos < bufLen
        nlPos = -1&
        bufIdx = pos
        Do While bufIdx < bufLen
            If buf(bufIdx) = 10 Then
                nlPos = bufIdx
                Exit Do
            End If
            bufIdx = bufIdx + 1
        Loop
        If nlPos < 0 Then Exit Do
        lineStart = pos
        lineBytes = nlPos - pos + 1
        bodyLen = nlPos - pos
        If bodyLen > 0 Then
            If buf(nlPos - 1) = 13 Then bodyLen = bodyLen - 1
        End If
        decodeOk = TqDecodeBody(buf, lineStart, bodyLen, bodyText)
        errKind = ""
        Set dupKeys = Nothing
        If Not decodeOk Then
            errKind = TQ_R_ENCODING
        Else
            validJson = TqJsonValidate(bodyText, startsObj, idOk, pidOk, dupKeys)
            If Not validJson Then
                errKind = "json_parse"
            ElseIf Not startsObj Then
                errKind = "not_an_object"
            Else
                If Not TqParseObject(bodyText, obj) Then
                    errKind = "json_parse"
                ElseIf Not idOk Then
                    errKind = "missing_id"
                ElseIf Not pidOk Then
                    errKind = "missing_pid"
                End If
            End If
        End If
        If Len(errKind) > 0 Then
            excerpt = TqExcerpt(decodeOk, bodyText, buf, lineStart, bodyLen)
            sampleHex = TqHexLower(buf, lineStart, bodyLen, 16&)
            Set badEntry = TqBadLine(seq + 1, baseOffset + lineStart, lineBytes, errKind, excerpt, sampleHex)
            If strictVal Then
                If errKind = TQ_R_ENCODING Then
                    strictReason = TQ_R_ENCODING
                Else
                    strictReason = TQ_R_BAD_JSONL
                End If
                Set strictBad = badEntry
                strictHit = True
                Exit Do
            End If
            seq = seq + 1
            scanned = scanned + 1
            badCount = badCount + 1
            If badLines.Count < TQ_MAX_BAD_LINES Then badLines.Add badEntry
            pos = nlPos + 1
        Else
            isMatch = True
            If Not IsNull(wId) Then
                If CStr(obj("id")) <> CStr(wId) Then isMatch = False
            End If
            If Not IsNull(wPrefix) Then
                idText = CStr(obj("id"))
                prefixText = CStr(wPrefix)
                If Left$(idText, Len(prefixText)) <> prefixText Then isMatch = False
            End If
            If Not IsNull(wPid) Then
                pidD = CDbl(obj("pid"))
                wantPid = CDbl(wPid)
                If pidD <> wantPid Then isMatch = False
            End If
            If isMatch And (Not countOnly) And (events.Count >= limitVal) Then
                truncatedBy = "limit"
                Exit Do
            End If
            seq = seq + 1
            scanned = scanned + 1
            If isMatch Then
                matched = matched + 1
                If firstMatchOffset < 0 Then firstMatchOffset = baseOffset + lineStart
                For Each kk In obj.Keys
                    fieldsSeen(CStr(kk)) = True
                Next kk
                If Not countOnly Then
                    Set fields = TqFieldsOf(obj, fieldFilter)
                    Set ev = CreateObject("Scripting.Dictionary")
                    ev("seq") = seq
                    ev("lineOffset") = baseOffset + lineStart
                    ev("lineBytes") = lineBytes
                    ev("id") = obj("id")
                    ev("pid") = obj("pid")
                    Set ev("fields") = fields
                    If includeRaw Then ev("raw") = bodyText
                    events.Add ev
                End If
            Else
                skipped = skipped + 1
            End If
            If Not dupKeys Is Nothing Then
                If dupKeys.Count > 0 Then
                    Set rec = CreateObject("Scripting.Dictionary")
                    rec("seq") = seq
                    Set rec("keys") = dupKeys
                    dupRecs.Add rec
                End If
            End If
            pos = nlPos + 1
        End If
    Loop
    consumed = pos
    nextOff = baseOffset + consumed
    reachedEof = False
    If IsNull(truncatedBy) Then
        If readHitEof Or (CDbl(baseOffset) + CDbl(bufLen) >= fileSize) Then
            reachedEof = True
        Else
            truncatedBy = "max_bytes"
        End If
    End If
    partial = bufLen - consumed
    If Not IsNull(truncatedBy) Then
        If CStr(truncatedBy) = "limit" Then partial = 0&
    End If
    If strictHit Then
        badCount = 1&
        Set badLines = New Collection
        badLines.Add strictBad
    End If
    Set res = CreateObject("Scripting.Dictionary")
    Set res("events") = events
    Set res("badLines") = badLines
    res("badCount") = badCount
    Set res("dupRecords") = dupRecs
    res("dupCount") = dupRecs.Count
    If dupRecs.Count > 0 Then
        Set rec = dupRecs(1)
        res("dupFirstSeq") = rec("seq")
    Else
        res("dupFirstSeq") = Null
    End If
    Set res("fieldsSeen") = fieldsSeen
    res("scanned") = scanned
    res("matched") = matched
    res("skipped") = skipped
    res("consumed") = consumed
    res("nextOffset") = nextOff
    res("partial") = partial
    res("reachedEof") = reachedEof
    res("truncatedBy") = truncatedBy
    res("strictReason") = strictReason
    If IsObject(strictBad) Then
        Set sbo = strictBad
        Set res("strictBad") = sbo
    Else
        res("strictBad") = Null
    End If
    Set TqScanWindow = res
End Function

' events[].fields: every top-level key except id/pid, projected by include_fields
' when one was given.  Object-valued fields are copied with Set (never a
' Variant-Let of an object, which JsonLib documents as a process crash).
Private Function TqFieldsOf(ByVal obj As Object, ByVal fieldFilter As Object) As Object
    Dim out As Object
    Dim k As Variant
    Dim keyStr As String
    Set out = CreateObject("Scripting.Dictionary")
    For Each k In obj.Keys
        keyStr = CStr(k)
        If keyStr <> "id" And keyStr <> "pid" Then
            If fieldFilter Is Nothing Then
                TqCopyItem out, obj, keyStr
            ElseIf fieldFilter.Exists(keyStr) Then
                TqCopyItem out, obj, keyStr
            End If
        End If
    Next k
    Set TqFieldsOf = out
End Function

Private Sub TqCopyItem(ByVal dst As Object, ByVal src As Object, ByVal key As String)
    Dim o As Object
    If IsObject(src(key)) Then
        Set o = src(key)
        Set dst(key) = o
    Else
        dst(key) = src(key)
    End If
End Sub

