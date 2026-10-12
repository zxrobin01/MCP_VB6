Attribute VB_Name = "Lessons"
Option Explicit
' Lessons (经验库) - modelled on the dynsim MCP mcp_lessons.py.
'
' Storage lives next to the add-in DLL:
'   mcp_lessons\defaults.json   shipped with the add-in (never written by tools)
'   mcp_lessons\user.json       user layer: overrides by id / disables / adds
'
' Lessons merge by id (user wins), filter on enabled, and are published to the
' client through initialize.instructions so an agent reads them before calling tools.

Private Const DIR_NAME As String = "mcp_lessons"

Public Function LessonsDir() As String
    Dim b As String
    b = App.Path
    If Right$(b, 1) <> "\" Then b = b & "\"
    LessonsDir = b & DIR_NAME
End Function

Public Function FilePath(ByVal leaf As String) As String
    FilePath = LessonsDir() & "\" & leaf
End Function

Public Sub EnsureDir()
    On Error Resume Next
    If Len(Dir$(LessonsDir(), vbDirectory)) = 0 Then MkDir LessonsDir()
End Sub

Private Function ReadText(ByVal path As String) As String
    Dim bb() As Byte, enc As String, s As String
    On Error Resume Next
    Encoding.ReadAllBytesTo path, bb
    If Encoding.ByteCount(bb) <= 0 Then Exit Function
    enc = Encoding.DetectEncoding(bb)
    Encoding.DecodeBytes bb, enc, s
    ReadText = s
End Function

Private Function WriteText(ByVal path As String, ByVal text As String) As Boolean
    ' TASK-P1-038: ATOMIC write - a half-written store must never be possible. The
    ' bytes go to <path>.tmp first, then the temp file is renamed over the target.
    Dim bb() As Byte, tmp As String
    On Error GoTo Bad
    EnsureDir
    Encoding.EncodeText text, "utf-8", bb
    tmp = path & ".tmp"
    Encoding.WriteAllBytesTo tmp, bb
    On Error Resume Next
    If Len(Dir$(path)) > 0 Then Kill path
    Err.Clear
    Name tmp As path
    If Err.Number <> 0 Then
        ' rename refused (locked/absent target): fall back to a direct write so the
        ' update is not lost, and report the outcome through the return value.
        Err.Clear
        Encoding.WriteAllBytesTo path, bb
        Err.Clear
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            WriteText = False
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0
    WriteText = True
    Exit Function
Bad:
    WriteText = False
End Function

Private Function Backup(ByVal path As String) As String
    Dim stamp As String, dest As String
    On Error Resume Next
    If Len(Dir$(path)) = 0 Then Exit Function
    stamp = Format$(Now, "yyyymmdd_hhnnss")
    dest = path & ".bak_" & stamp
    FileCopy path, dest
    Backup = dest
End Function

Public Function LoadDoc(ByVal path As String) As Object
    Dim txt As String, v As Variant
    txt = ReadText(path)
    If Len(txt) = 0 Then Exit Function
    On Error Resume Next
    If JsonParse(txt, v) Then
        If IsObject(v) Then
            If TypeName(v) = "Dictionary" Then Set LoadDoc = v
        End If
    End If
End Function

' ---------- merging ----------

Private Function CloneDict(ByVal src As Object) As Object
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In src.Keys
        If IsObject(src(k)) Then
            Set d(CStr(k)) = src(k)
        Else
            d(CStr(k)) = src(k)
        End If
    Next
    Set CloneDict = d
End Function

Private Sub MergeInto(ByVal dst As Object, ByVal src As Object)
    Dim k As Variant
    For Each k In src.Keys
        If LCase$(CStr(k)) <> "id" Then
            If IsObject(src(k)) Then
                Set dst(CStr(k)) = src(k)
            Else
                dst(CStr(k)) = src(k)
            End If
        End If
    Next
End Sub

Private Sub MergeList(ByVal doc As Object, ByRef order As Collection, ByRef byId As Object)
    If doc Is Nothing Then Exit Sub
    Dim arr As Variant, isObj As Boolean, i As Long, it As Object, k As String, ex As Object
    If Not JsonGet(doc, "lessons", arr, isObj) Then Exit Sub
    If Not isObj Then Exit Sub
    If TypeName(arr) <> "Collection" Then Exit Sub
    For i = 1 To arr.Count
        MaybePump
        If IsObject(arr(i)) Then
            Set it = arr(i)
            If TypeName(it) = "Dictionary" Then
                k = LCase$(JsonGetStr(it, "id", ""))
                If Len(k) > 0 Then
                    If byId.Exists(k) Then
                        Set ex = byId(k)
                        MergeInto ex, it
                    Else
                        byId.Add k, CloneDict(it)
                        order.Add k
                    End If
                End If
            End If
        End If
    Next
End Sub

Public Function MergedLessons(ByVal includeDisabled As Boolean) As Collection
    Dim order As New Collection, byId As Object, out As New Collection
    Dim i As Long, it As Object
    Set byId = CreateObject("Scripting.Dictionary")
    MergeList LoadDoc(FilePath("defaults.json")), order, byId
    MergeList LoadDoc(FilePath("user.json")), order, byId
    For i = 1 To order.Count
        MaybePump
        Set it = byId(order(i))
        If includeDisabled Then
            out.Add it
        ElseIf JsonGetBool(it, "enabled", True) Then
            out.Add it
        End If
    Next
    Set MergedLessons = out
End Function

Private Function RulesText(ByVal source As Object) As String
    Dim rules As Variant, isObj As Boolean, i As Long, s As String
    If source Is Nothing Then Exit Function
    If Not JsonGet(source, "rules", rules, isObj) Then Exit Function
    If isObj Then
        If TypeName(rules) = "Collection" Then
            For i = 1 To rules.Count
                MaybePump
                If i > 1 Then s = s & vbCrLf
                s = s & CStr(i) & ") " & CStr(rules(i))
            Next
        End If
    ElseIf Not IsNull(rules) Then
        s = CStr(rules)
    End If
    RulesText = s
End Function

Public Function Instructions() As String
    Dim doc As Object, inst As Variant, isObj As Boolean, intro As String, body As String
    intro = "VB6MCP usage notes (read before calling tools):"
    Set doc = LoadDoc(FilePath("defaults.json"))
    If Not doc Is Nothing Then
        If JsonGet(doc, "instructions", inst, isObj) Then
            If isObj Then
                If TypeName(inst) = "Dictionary" Then
                    intro = JsonGetStr(inst, "intro", intro)
                    body = RulesText(inst)
                End If
            End If
        End If
    End If
    If Len(body) > 0 Then Instructions = intro & vbCrLf & body Else Instructions = intro
End Function
' ---------- JSON API used by the MCP tools ----------

Private Function Q(ByVal s As String) As String
    Q = Chr$(34) & s & Chr$(34)
End Function

Private Function JStr(ByVal s As String) As String
    Dim t As String
    t = Replace(s, Chr$(92), Chr$(92) & Chr$(92))
    t = Replace(t, Chr$(34), Chr$(92) & Chr$(34))
    t = Replace(t, vbCrLf, Chr$(92) & "n")
    t = Replace(t, vbCr, Chr$(92) & "n")
    t = Replace(t, vbLf, Chr$(92) & "n")
    t = Replace(t, vbTab, Chr$(92) & "t")
    JStr = Chr$(34) & t & Chr$(34)
End Function

Private Function NextCustomId(ByVal lessons As Collection) As String
    Dim n As Long, i As Long, cand As String, used As Boolean
    n = 1
    Do While n < 1000
        cand = "custom-" & CStr(n)
        used = False
        For i = 1 To lessons.Count
            If IsObject(lessons(i)) Then
                If LCase$(JsonGetStr(lessons(i), "id", "")) = LCase$(cand) Then used = True
            End If
        Next
        If Not used Then
            NextCustomId = cand
            Exit Function
        End If
        n = n + 1
    Loop
    NextCustomId = "custom-x"
End Function

Private Function UserDoc() As Object
    Dim doc As Object
    Set doc = LoadDoc(FilePath("user.json"))
    If doc Is Nothing Then Set doc = CreateObject("Scripting.Dictionary")
    Set UserDoc = doc
End Function

' ---------------------------------------------------------------------------
' TASK-P1-038 (REQ-20260921-005): a store that EXISTS but does not parse must
' NEVER be silently replaced by an empty one. UserDoc() above returns an empty
' dictionary in that case, and the writers used to persist it - which wiped the
' user layer (silent data loss). The writers now go through this strict loader:
'   missing file / whitespace-only file -> a fresh store (nothing to lose)
'   present but unparseable, or not a JSON object -> NOTHING + a timestamped .bak
' ---------------------------------------------------------------------------
Private Function LoadUserDocStrict(ByRef why As String, ByRef backupPath As String) As Object
    Dim path As String, txt As String, v As Variant, doc As Object
    path = FilePath("user.json")
    why = ""
    backupPath = ""
    txt = ReadText(path)
    If Len(Dir$(path)) = 0 Then
        why = "missing"
        Set LoadUserDocStrict = CreateObject("Scripting.Dictionary")
        Exit Function
    End If
    If Len(Trim$(txt)) = 0 Then
        ' TASK-P1-038: an EXISTING but empty store is refused, not treated as fresh. A
        ' zero-byte file is indistinguishable from a store that a previous truncating
        ' write emptied, so the safe reading wins: refuse, keep the file (and a .bak),
        ' and tell the caller how to start a fresh store.
        why = "empty"
        backupPath = Backup(path)
        Exit Function
    End If
    On Error Resume Next
    If JsonParse(txt, v) Then
        If IsObject(v) Then
            If TypeName(v) = "Dictionary" Then Set doc = v
        End If
    End If
    Err.Clear
    On Error GoTo 0
    If doc Is Nothing Then
        why = "unreadable"
        backupPath = Backup(path)
        Exit Function
    End If
    Set LoadUserDocStrict = doc
End Function

Private Function StoreUnreadableJson(ByVal why As String, ByVal backupPath As String) As String
    Dim s As String
    s = "{" & Q("ok") & ":false," & Q("reason") & ":" & JStr("lesson_store_unreadable")
    s = s & "," & Q("path") & ":" & JStr(FilePath("user.json"))
    s = s & "," & Q("backupPath") & ":" & JStr(backupPath)
    s = s & "," & Q("error") & ":" & JStr("the lesson store exists but could not be parsed (" & why & "); it was left UNCHANGED")
    s = s & "," & Q("advice") & ":" & JStr("repair or remove " & FilePath("user.json") & " (a timestamped backup was kept when possible), then retry; nothing was written")
    StoreUnreadableJson = s & "}"
End Function

Private Function UserLessons(ByVal doc As Object) As Collection
    Dim c As Variant, isObj As Boolean
    If JsonGet(doc, "lessons", c, isObj) Then
        If isObj Then
            If TypeName(c) = "Collection" Then
                Set UserLessons = c
                Exit Function
            End If
        End If
    End If
    Dim nc As New Collection
    Set doc("lessons") = nc
    Set UserLessons = nc
End Function

Public Function ListJson(ByVal includeDisabled As Boolean) As String
    Dim col As Collection, i As Long, it As Object, s As String
    Set col = MergedLessons(includeDisabled)
    For i = 1 To col.Count
        MaybePump
        Set it = col(i)
        If Len(s) > 0 Then s = s & ","
        s = s & "{" & Q("id") & ":" & JStr(JsonGetStr(it, "id", "")) _
            & "," & Q("scenario") & ":" & JStr(JsonGetStr(it, "scenario", "")) _
            & "," & Q("trap") & ":" & JStr(JsonGetStr(it, "trap", "")) _
            & "," & Q("do") & ":" & JStr(JsonGetStr(it, "do", "")) _
            & "," & Q("scope") & ":" & JStr(JsonGetStr(it, "scope", "all")) _
            & "," & Q("enabled") & ":" & LCase$(CStr(JsonGetBool(it, "enabled", True))) & "}"
    Next
    ListJson = "{" & Q("count") & ":" & CStr(col.Count) & "," & Q("lessons") & ":[" & s & "]}"
End Function

Public Function AddJson(ByVal id As String, ByVal scenario As String, ByVal trapText As String, ByVal doText As String, ByVal scope As String, ByVal enabled As Boolean) As String
    If Len(scenario) = 0 Or Len(trapText) = 0 Or Len(doText) = 0 Then
        AddJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("scenario, trap and do are all required") & "}"
        Exit Function
    End If
    Dim doc As Object, lessons As Collection, i As Long, action As String, item As Object
    Dim why As String, bkp As String
    Set doc = LoadUserDocStrict(why, bkp)
    If doc Is Nothing Then
        AddJson = StoreUnreadableJson(why, bkp)
        Exit Function
    End If
    Set lessons = UserLessons(doc)
    If Len(id) = 0 Then id = NextCustomId(lessons)
    Set item = CreateObject("Scripting.Dictionary")
    item("id") = id
    If Len(scope) = 0 Then item("scope") = "all" Else item("scope") = scope
    item("enabled") = enabled
    item("scenario") = scenario
    item("trap") = trapText
    item("do") = doText
    action = "added"
    For i = 1 To lessons.Count
        MaybePump
        If IsObject(lessons(i)) Then
            If LCase$(JsonGetStr(lessons(i), "id", "")) = LCase$(id) Then
                lessons.Remove i
                lessons.Add item, , i
                action = "updated"
                Exit For
            End If
        End If
    Next
    If action = "added" Then lessons.Add item
    Backup FilePath("user.json")
    If Not WriteText(FilePath("user.json"), JsonStringify(doc)) Then
        AddJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("cannot write user.json") & "}"
        Exit Function
    End If
    AddJson = "{" & Q("ok") & ":true," & Q("id") & ":" & JStr(id) & "," & Q("action") & ":" & JStr(action) & "," & Q("lessonsTotal") & ":" & CStr(MergedLessons(False).Count) & "}"
End Function

Private Function InDefaults(ByVal id As String) As Boolean
    Dim dDoc As Object, arr As Variant, isObj As Boolean, j As Long
    Set dDoc = LoadDoc(FilePath("defaults.json"))
    If dDoc Is Nothing Then Exit Function
    If Not JsonGet(dDoc, "lessons", arr, isObj) Then Exit Function
    If Not isObj Then Exit Function
    If TypeName(arr) <> "Collection" Then Exit Function
    For j = 1 To arr.Count
        If IsObject(arr(j)) Then
            If LCase$(JsonGetStr(arr(j), "id", "")) = LCase$(id) Then
                InDefaults = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function DeleteJson(ByVal id As String) As String
    If Len(id) = 0 Then
        DeleteJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("id is required") & "}"
        Exit Function
    End If
    Dim doc As Object, lessons As Collection, i As Long, action As String, removed As Boolean
    Dim why As String, bkp As String
    Set doc = LoadUserDocStrict(why, bkp)
    If doc Is Nothing Then
        DeleteJson = StoreUnreadableJson(why, bkp)
        Exit Function
    End If
    Set lessons = UserLessons(doc)
    removed = False
    For i = lessons.Count To 1 Step -1
        MaybePump
        If IsObject(lessons(i)) Then
            If LCase$(JsonGetStr(lessons(i), "id", "")) = LCase$(id) Then
                lessons.Remove i
                removed = True
            End If
        End If
    Next
    If InDefaults(id) Then
        Dim dis As Object
        Set dis = CreateObject("Scripting.Dictionary")
        dis("id") = id
        dis("enabled") = False
        lessons.Add dis
        action = "disabled(default)"
    ElseIf removed Then
        action = "removed"
    Else
        action = "not_found"
    End If
    Backup FilePath("user.json")
    If Not WriteText(FilePath("user.json"), JsonStringify(doc)) Then
        DeleteJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("cannot write user.json") & "}"
        Exit Function
    End If
    DeleteJson = "{" & Q("ok") & ":" & LCase$(CStr(action <> "not_found")) & "," & Q("id") & ":" & JStr(id) & "," & Q("action") & ":" & JStr(action) & "," & Q("lessonsTotal") & ":" & CStr(MergedLessons(False).Count) & "}"
End Function

Public Function ExportMd(ByVal includeDisabled As Boolean) As String
    Dim col As Collection, i As Long, s As String, it As Object
    Set col = MergedLessons(includeDisabled)
    s = "# VB6MCP lessons (re-import with import_lessons_md)" & vbCrLf
    For i = 1 To col.Count
        MaybePump
        Set it = col(i)
        s = s & vbCrLf & "## [" & JsonGetStr(it, "id", "") & "] " & JsonGetStr(it, "scenario", "") & vbCrLf
        s = s & ChrW$(&H5751) & ": " & JsonGetStr(it, "trap", "") & vbCrLf
        s = s & ChrW$(&H5E94) & ": " & JsonGetStr(it, "do", "") & vbCrLf
        s = s & "scope: " & JsonGetStr(it, "scope", "all") & vbCrLf
        If Not JsonGetBool(it, "enabled", True) Then s = s & "enabled: false" & vbCrLf
    Next
    ExportMd = s
End Function

' ---------- import: '## [id]' handling ----------

' '## [id] scenario' -> id + scenario text. '## scenario' -> no id, scenario as written.
' A malformed bracket (no ']'), a blank id, or an id with whitespace/brackets all leave
' idOut empty so AddJson mints custom-N: an existing lesson is never overwritten.
Private Sub SplitMdHead(ByVal rest As String, ByRef idOut As String, ByRef scenOut As String)
    Dim q As Long, cand As String
    idOut = ""
    scenOut = rest
    If Left$(rest, 1) <> "[" Then Exit Sub
    q = InStr(rest, "]")
    If q < 2 Then Exit Sub
    cand = Trim$(Mid$(rest, 2, q - 2))
    If ValidLessonId(cand) Then idOut = cand
    scenOut = Trim$(Mid$(rest, q + 1))
    If Len(scenOut) = 0 Then
        If Len(idOut) > 0 Then
            scenOut = idOut
        Else
            scenOut = rest
        End If
    End If
End Sub

Private Function ValidLessonId(ByVal s As String) As Boolean
    Dim i As Long, c As Long
    If Len(s) = 0 Then Exit Function
    If Len(s) > 120 Then Exit Function
    For i = 1 To Len(s)
        c = AscW(Mid$(s, i, 1))
        If c >= 1 And c <= 32 Then Exit Function
        If c = 91 Or c = 93 Then Exit Function
    Next
    ValidLessonId = True
End Function

' Ids already present in the MERGED library (defaults + user, disabled included), lowercased.
Private Function MergedIdSet() As Object
    Dim d As Object, col As Collection, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    Set col = MergedLessons(True)
    For i = 1 To col.Count
        k = LCase$(JsonGetStr(col(i), "id", ""))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d.Add k, True
        End If
    Next
    Set MergedIdSet = d
End Function

Private Function ParseMd(ByVal text As String, ByRef outIds As Collection, ByRef outScen As Collection, ByRef outTrap As Collection, ByRef outDo As Collection) As Long
    Dim lines() As String, i As Long, cur As Long, ln As String, rest As String
    Dim headId As String, headScen As String
    Set outIds = New Collection
    Set outScen = New Collection
    Set outTrap = New Collection
    Set outDo = New Collection
    lines = Split(Replace(text, vbCrLf, vbLf), vbLf)
    cur = -1
    For i = 0 To UBound(lines)
        MaybePump
        ln = Trim$(lines(i))
        If Left$(ln, 2) = "##" Then
            rest = Trim$(Mid$(ln, 3))
            ' '## [id] scenario' carries the id; '## scenario' has none (custom-N fallback).
            SplitMdHead rest, headId, headScen
            outIds.Add headId
            outScen.Add headScen
            outTrap.Add ""
            outDo.Add ""
            cur = outIds.Count
        ElseIf cur > 0 Then
            If Left$(ln, 2) = ChrW$(&H5751) & ":" Or LCase$(Left$(ln, 4)) = "trap" Then
                ' Collection.Item is read-only (coll(cur) = v raises 424), and
                ' Collection.Add before:=cur raises 5 while the collection is still
                ' empty. cur is always the most recent '##' header, i.e. the LAST
                ' element of every collection, so replace the last item instead:
                ' order and the 4-way index alignment are both preserved.
                outTrap.Remove outTrap.Count
                outTrap.Add Trim$(Mid$(ln, InStr(ln, ":") + 1))
            ElseIf Left$(ln, 2) = ChrW$(&H5E94) & ":" Or LCase$(Left$(ln, 2)) = "do" Then
                outDo.Remove outDo.Count
                outDo.Add Trim$(Mid$(ln, InStr(ln, ":") + 1))
            End If
        End If
    Next
    ParseMd = outIds.Count
End Function

Public Function ImportMdJson(ByVal path As String, ByVal dryRun As Boolean) As String
    Dim text As String, ids As Collection, scen As Collection, tr As Collection, ac As Collection
    Dim n As Long, i As Long, added As Long, skipped As Long
    Dim conflicts As Long, idsUsed As Long, seen As Object, wantId As String, useId As String
    text = ReadText(path)
    ' A tool must never leak an exception: the smoke run saw this surface as
    ' "Error: 需要对象" (VB 424) instead of a clean ok:false.
    On Error GoTo Bad
    If Len(text) = 0 Then
        ImportMdJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("cannot read: " & path) & "}"
        Exit Function
    End If
    n = ParseMd(text, ids, scen, tr, ac)
    If n = 0 Then
        ImportMdJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("no lessons parsed: expected '## [id] scenario' headers plus trap/do lines (run export_lessons to see the exact format)") & "}"
        Exit Function
    End If
    ' Conflict rule (TASK-P1-002): an id already present in the MERGED library is NOT
    ' overwritten and NOT merged -- the incoming lesson falls back to custom-N. The
    ' analysis also runs for dryRun, so conflicts/idsUsed report what WOULD happen.
    Set seen = MergedIdSet()
    For i = 1 To n
        MaybePump
        If Len(CStr(tr(i))) > 0 And Len(CStr(ac(i))) > 0 Then
            wantId = CStr(ids(i))
            useId = ""
            If Len(wantId) > 0 Then
                If seen.Exists(LCase$(wantId)) Then
                    conflicts = conflicts + 1
                Else
                    useId = wantId
                    seen.Add LCase$(wantId), True
                    idsUsed = idsUsed + 1
                End If
            End If
            If Not dryRun Then
                ' TASK-P1-038: AddJson refuses when the store is unreadable; do not
                ' count that as added (a false "added" would be a dishonest receipt).
                Dim rj As String
                rj = AddJson(useId, CStr(scen(i)), CStr(tr(i)), CStr(ac(i)), "all", True)
                If InStr(rj, Q("lesson_store_unreadable")) > 0 Then
                    ImportMdJson = rj
                    Exit Function
                End If
                added = added + 1
            End If
        Else
            skipped = skipped + 1
        End If
    Next
    ImportMdJson = "{" & Q("ok") & ":true," & Q("parsed") & ":" & CStr(n) & "," & Q("added") & ":" & CStr(added) & "," & Q("skipped") & ":" & CStr(skipped) & "," & Q("idsUsed") & ":" & CStr(idsUsed) & "," & Q("conflicts") & ":" & CStr(conflicts) & "," & Q("dryRun") & ":" & LCase$(CStr(dryRun)) & "}"
    Exit Function
Bad:
    ImportMdJson = "{" & Q("ok") & ":false," & Q("error") & ":" & JStr("import failed: " & CStr(Err.Number) & " " & Err.Description) & "}"
End Function
