Attribute VB_Name = "modWordUtil"
Option Explicit

' Shared Word helpers: reading visible paragraph text (tracked deletions removed)
' with a mapping back to document positions, tool comments, revision view.

Private mBatchActive As Boolean
Private mBatchOldName As String
Private mBatchOldInitials As String

' state of the running command (BeginDocWork / EndDocWork)
Private mWorkDoc As Document
Private mWorkView As Variant
Private mWorkTrack As Boolean
Private mWorkUndo As Boolean

' Tool comments carry the author from settings (check.commentAuthor), so that
' "Pruefmarkierungen entfernen" can later remove exactly these comments.
' For many comments call BeginCommentBatch/EndCommentBatch around them, so the
' Word user name is switched only once.
Public Sub BeginCommentBatch()
    If mBatchActive Then Exit Sub
    mBatchOldName = Application.UserName
    mBatchOldInitials = Application.UserInitials
    Application.UserName = ToolAuthor()
    Application.UserInitials = "PC"
    mBatchActive = True
End Sub

Public Sub EndCommentBatch()
    If Not mBatchActive Then Exit Sub
    Application.UserName = mBatchOldName
    Application.UserInitials = mBatchOldInitials
    mBatchActive = False
End Sub

Public Function ToolAuthor() As String
    ToolAuthor = CStr(GetSettingValue("check.commentAuthor", "PatentCheck"))
End Function

Public Sub AddToolComment(ByVal doc As Document, ByVal startPos As Long, ByVal endPos As Long, ByVal commentText As String)
    Dim c As Comment, errNum As Long, errDesc As String, ownBatch As Boolean
    If Not mBatchActive Then
        BeginCommentBatch
        ownBatch = True
    End If

    On Error Resume Next
    Set c = doc.Comments.Add(doc.Range(startPos, endPos), commentText)
    errNum = Err.Number
    errDesc = Err.Description
    If errNum = 0 Then
        c.Author = ToolAuthor()
        c.Initial = "PC"
    End If
    Err.Clear
    On Error GoTo 0

    If ownBatch Then EndCommentBatch
    If errNum <> 0 Then Err.Raise errNum, "AddToolComment", errDesc
End Sub

' True for comments created by Patent-Tools: author from settings, or text
' starting with "[<check name>] ".
Public Function IsToolComment(ByVal c As Comment) As Boolean
    Dim t As String, q As Long, keys As Variant, i As Long
    If c.Author = ToolAuthor() Then
        IsToolComment = True
        Exit Function
    End If
    t = c.Range.Text
    If Left$(t, 1) <> "[" Then Exit Function
    q = InStr(t, "] ")
    If q = 0 Then Exit Function
    keys = Split(CHECK_KEYS & "|renumber", "|")
    For i = LBound(keys) To UBound(keys)
        If Mid$(t, 2, q - 2) = CheckDisplayName(CStr(keys(i))) Then
            IsToolComment = True
            Exit Function
        End If
    Next i
End Function

' Removes all tool comments and the highlight of their scope. Returns the count.
Public Function RemoveToolMarks(ByVal doc As Document) As Long
    Dim i As Long, c As Comment
    For i = doc.Comments.Count To 1 Step -1
        Set c = doc.Comments(i)
        If IsToolComment(c) Then
            If c.Scope.End > c.Scope.Start Then c.Scope.HighlightColorIndex = wdNoHighlight
            c.Delete
            RemoveToolMarks = RemoveToolMarks + 1
        End If
    Next i
End Function

Public Function HighlightIndexFromName(ByVal colorName As String) As Long
    Select Case LCase$(colorName)
        Case "wdyellow": HighlightIndexFromName = wdYellow
        Case "wdbrightgreen": HighlightIndexFromName = wdBrightGreen
        Case "wdturquoise": HighlightIndexFromName = wdTurquoise
        Case "wdpink": HighlightIndexFromName = wdPink
        Case "wdblue": HighlightIndexFromName = wdBlue
        Case "wdred": HighlightIndexFromName = wdRed
        Case "wddarkblue": HighlightIndexFromName = wdDarkBlue
        Case "wdteal": HighlightIndexFromName = wdTeal
        Case "wdgreen": HighlightIndexFromName = wdGreen
        Case "wdviolet": HighlightIndexFromName = wdViolet
        Case "wddarkred": HighlightIndexFromName = wdDarkRed
        Case "wddarkyellow": HighlightIndexFromName = wdDarkYellow
        Case "wdgray50": HighlightIndexFromName = wdGray50
        Case "wdgray25": HighlightIndexFromName = wdGray25
        Case "none", "wdnohighlight", "": HighlightIndexFromName = wdNoHighlight
        Case Else: HighlightIndexFromName = wdYellow
    End Select
End Function

' Shows all mark-up inline so Range.Text contains deleted text too and its
' positions match the document. Returns the previous state for RestoreRevisionView.
Public Function ForceInlineMarkupView(ByVal doc As Document) As Variant
    Dim v As View
    On Error Resume Next
    Set v = doc.ActiveWindow.View
    If v Is Nothing Then Exit Function
    ForceInlineMarkupView = Array(v.ShowRevisionsAndComments, v.RevisionsFilter.Markup, _
        v.RevisionsFilter.View, v.MarkupMode)
    v.ShowRevisionsAndComments = True
    v.RevisionsFilter.Markup = wdRevisionsMarkupAll
    v.RevisionsFilter.View = wdRevisionsViewFinal
    v.MarkupMode = wdInLineRevisions
End Function

Public Sub RestoreRevisionView(ByVal doc As Document, ByVal state As Variant)
    Dim v As View
    If IsEmpty(state) Then Exit Sub
    On Error Resume Next
    Set v = doc.ActiveWindow.View
    v.ShowRevisionsAndComments = state(0)
    v.RevisionsFilter.Markup = state(1)
    v.RevisionsFilter.View = state(2)
    v.MarkupMode = state(3)
End Sub

' ---------------------------------------------------------------- reading visible text

' Reads every paragraph of the main text. For each paragraph i:
'   texts(i)  visible text (tracked deletions removed, no paragraph mark)
'   starts(i) document position of the paragraph start
'   segs(i)   Empty if visible text maps 1:1 onto the document, otherwise an
'             array of segments Array(visibleStart, docStart, length)
'   indents(i) left indent in points (used to recognise sub-feature groups)
Public Sub ReadVisibleParagraphs(ByVal doc As Document, ByRef texts As Variant, _
        ByRef starts As Variant, ByRef segs As Variant, Optional ByRef indents As Variant)
    Dim delStart() As Long, delEnd() As Long, delCount As Long
    Dim t() As String, st() As Long, sg() As Variant, ind() As Single
    Dim para As Paragraph, rng As Range, raw As String
    Dim i As Long, rs As Long, re As Long, ptr As Long

    CollectDeletions doc, delStart, delEnd, delCount
    ReDim t(0 To doc.Paragraphs.Count - 1)
    ReDim st(0 To doc.Paragraphs.Count - 1)
    ReDim sg(0 To doc.Paragraphs.Count - 1)
    ReDim ind(0 To doc.Paragraphs.Count - 1)

    For Each para In doc.Paragraphs
        Set rng = para.Range
        raw = rng.Text
        If Right$(raw, 1) = vbCr Then raw = Left$(raw, Len(raw) - 1)
        rs = rng.Start
        re = rs + Len(raw)
        st(i) = rs
        ind(i) = para.LeftIndent

        ' deletions are sorted; skip those that end before this paragraph
        Do While ptr < delCount
            If delEnd(ptr) > rs Then Exit Do
            ptr = ptr + 1
        Loop
        If ptr < delCount Then
            If delStart(ptr) < re Then
                sg(i) = VisibleSegments(raw, rs, re, delStart, delEnd, delCount, ptr, t(i))
            Else
                t(i) = raw
            End If
        Else
            t(i) = raw
        End If
        i = i + 1
    Next para

    texts = t
    starts = st
    segs = sg
    indents = ind
End Sub

Private Function VisibleSegments(ByVal raw As String, ByVal rs As Long, ByVal re As Long, _
        delStart() As Long, delEnd() As Long, ByVal delCount As Long, ByVal ptr As Long, _
        ByRef visible As String) As Variant
    Dim segs As New Collection, arr() As Variant, k As Long, cur As Long
    Dim ds As Long, de As Long, visPos As Long
    cur = rs
    visPos = 1
    visible = ""
    k = ptr
    Do While k < delCount
        If delStart(k) >= re Then Exit Do
        ds = delStart(k)
        If ds < cur Then ds = cur
        de = delEnd(k)
        If de > re Then de = re
        If ds > cur Then
            segs.Add Array(visPos, cur, ds - cur)
            visible = visible & Mid$(raw, cur - rs + 1, ds - cur)
            visPos = visPos + (ds - cur)
        End If
        If de > cur Then cur = de
        k = k + 1
    Loop
    If re > cur Then
        segs.Add Array(visPos, cur, re - cur)
        visible = visible & Mid$(raw, cur - rs + 1, re - cur)
    End If
    If segs.Count = 0 Then
        VisibleSegments = Array()
        Exit Function
    End If
    ReDim arr(0 To segs.Count - 1)
    For k = 1 To segs.Count
        arr(k - 1) = segs(k)
    Next k
    VisibleSegments = arr
End Function

' Collects tracked deletions (incl. moved-away text) sorted by start position.
Private Sub CollectDeletions(ByVal doc As Document, ByRef delStart() As Long, _
        ByRef delEnd() As Long, ByRef delCount As Long)
    Dim rev As Revision, n As Long, i As Long, j As Long, s As Long, e As Long
    n = doc.Content.Revisions.Count
    ReDim delStart(0 To n)
    ReDim delEnd(0 To n)
    delCount = 0
    If n = 0 Then Exit Sub
    For Each rev In doc.Content.Revisions
        If rev.Type = wdRevisionDelete Or rev.Type = wdRevisionMovedFrom Then
            delStart(delCount) = rev.Range.Start
            delEnd(delCount) = rev.Range.End
            delCount = delCount + 1
        End If
    Next rev
    For i = 1 To delCount - 1
        s = delStart(i)
        e = delEnd(i)
        j = i - 1
        Do While j >= 0
            If delStart(j) <= s Then Exit Do
            delStart(j + 1) = delStart(j)
            delEnd(j + 1) = delEnd(j)
            j = j - 1
        Loop
        delStart(j + 1) = s
        delEnd(j + 1) = e
    Next i
End Sub

' Maps a visible span of paragraph p to document positions [a, b).
' Fails if the span is interrupted by tracked deletions.
Public Function MapSpan(ByVal starts As Variant, ByVal segs As Variant, ByVal p As Long, _
        ByVal visStart As Long, ByVal visLen As Long, ByRef a As Long, ByRef b As Long) As Boolean
    a = DocPos(starts(p), segs(p), visStart)
    b = DocPos(starts(p), segs(p), visStart + visLen - 1)
    If a < 0 Or b < 0 Then Exit Function
    b = b + 1
    MapSpan = (b - a = visLen)
End Function

Private Function DocPos(ByVal paraStart As Long, ByVal seg As Variant, ByVal visPos As Long) As Long
    Dim k As Long, s As Variant
    If IsEmpty(seg) Then
        DocPos = paraStart + visPos - 1
        Exit Function
    End If
    DocPos = -1
    For k = LBound(seg) To UBound(seg)
        s = seg(k)
        If visPos >= s(0) And visPos < s(0) + s(2) Then
            DocPos = s(1) + visPos - s(0)
            Exit Function
        End If
    Next k
End Function

' ---------------------------------------------------------------- applying text operations

' Applies replace/insert operations (OP_* records; length 0 = insert) from the
' end backwards. Each replacement is verified against the expected old text.
' Returns the number applied; failed counts those that could not be applied.
Public Function ApplyTextOps(ByVal doc As Document, ByVal texts As Variant, ByVal starts As Variant, _
        ByVal segs As Variant, ByVal ops As Collection, ByRef failed As Long) As Long
    Dim sorted As Variant, i As Long, op As Variant
    sorted = SortOps(ops)
    failed = 0
    For i = UBound(sorted) To LBound(sorted) Step -1
        op = sorted(i)
        If op(OP_KIND) <> OPK_FLAG Then
            If ApplyReplaceOp(doc, texts, starts, segs, op) Then
                ApplyTextOps = ApplyTextOps + 1
            Else
                failed = failed + 1
            End If
        End If
    Next i
End Function

Public Function ApplyReplaceOp(ByVal doc As Document, ByVal texts As Variant, ByVal starts As Variant, _
        ByVal segs As Variant, ByVal op As Variant) As Boolean
    Dim a As Long, b As Long, rng As Range, p As Long, visLen As Long
    p = op(OP_PARA)
    If op(OP_LEN) = 0 Then
        visLen = Len(texts(p))
        If op(OP_START) <= visLen Then
            a = DocPos(starts(p), segs(p), op(OP_START))
        ElseIf visLen > 0 Then
            a = DocPos(starts(p), segs(p), visLen)
            If a >= 0 Then a = a + 1
        Else
            a = starts(p)
        End If
        If a < 0 Then Exit Function
        doc.Range(a, a).InsertAfter op(OP_TEXT)
        ApplyReplaceOp = True
        Exit Function
    End If
    If Not MapSpan(starts, segs, p, op(OP_START), op(OP_LEN), a, b) Then Exit Function
    Set rng = doc.Range(a, b)
    If rng.Text <> Mid$(texts(p), op(OP_START), op(OP_LEN)) Then Exit Function
    rng.Text = op(OP_TEXT)
    ApplyReplaceOp = True
End Function

' ---------------------------------------------------------------- document work session

' BeginDocWork/EndDocWork wrap a command: settings reload, screen updating off,
' inline mark-up view, one undo step; EndDocWork restores everything (also
' after errors). SetTrackMode "silent" switches Track Changes off meanwhile.

Public Function BeginDocWork(ByVal undoTitle As String) As Document
    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Function
    End If
    Set mWorkDoc = ActiveDocument
    ReloadSettings
    Application.ScreenUpdating = False
    mWorkTrack = mWorkDoc.TrackRevisions
    mWorkView = ForceInlineMarkupView(mWorkDoc)
    If Len(undoTitle) > 0 Then
        Application.UndoRecord.StartCustomRecord undoTitle
        mWorkUndo = True
    End If
    Set BeginDocWork = mWorkDoc
End Function

Public Sub SetTrackMode(ByVal mode As String)
    If mWorkDoc Is Nothing Then Exit Sub
    If LCase$(mode) = "silent" Then
        mWorkDoc.TrackRevisions = False
    Else
        mWorkDoc.TrackRevisions = mWorkTrack
    End If
End Sub

Public Sub EndDocWork()
    On Error Resume Next
    EndCommentBatch
    If Not mWorkDoc Is Nothing Then
        mWorkDoc.TrackRevisions = mWorkTrack
        RestoreRevisionView mWorkDoc, mWorkView
    End If
    If mWorkUndo Then Application.UndoRecord.EndCustomRecord
    mWorkUndo = False
    Set mWorkDoc = Nothing
    Application.ScreenUpdating = True
End Sub
