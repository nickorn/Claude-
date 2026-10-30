Attribute VB_Name = "modRenumber"
Option Explicit

' MVP-1: renumber claims in document order and update all back-references.
'
' Planning happens on plain text (modClaimParse.PlanRenumber). This module reads
' the visible paragraph texts, maps planned edits back to document positions and
' applies them from the end of the document backwards, so earlier positions stay
' valid. Tracked deletions are excluded from the visible text, so claim sets that
' already contain mark-up are handled correctly.
'
' Track Changes (settings: renumber.trackChangesMode):
'   "tracked" - changes follow the document's Track Changes state (recorded if on)
'   "silent"  - Track Changes is switched off while renumbering

Public Sub RenumberClaims()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant
    Dim ops As New Collection, sorted As Variant, op As Variant
    Dim errMsg As String, claimCount As Long, i As Long
    Dim labelsChanged As Long, refsChanged As Long, flagCount As Long
    Dim viewState As Variant, oldTrack As Boolean, undoOpen As Boolean, stateSaved As Boolean
    Dim errDesc As String

    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Sub
    End If
    Set doc = ActiveDocument

    On Error GoTo Fail
    ReloadSettings
    Application.ScreenUpdating = False
    oldTrack = doc.TrackRevisions
    viewState = ForceInlineMarkupView(doc)
    stateSaved = True

    ReadVisibleParagraphs doc, texts, starts, segs
    If Not PlanRenumber(texts, ops, errMsg, claimCount) Then
        RestoreState doc, viewState, oldTrack, undoOpen
        MsgBox errMsg, vbExclamation, AppTitle()
        Exit Sub
    End If
    If ops.Count = 0 Then
        RestoreState doc, viewState, oldTrack, undoOpen
        MsgBox MsgAlreadyConsecutive(claimCount), vbInformation, AppTitle()
        Exit Sub
    End If

    Application.UndoRecord.StartCustomRecord UndoTitleRenumber()
    undoOpen = True
    If LCase$(CStr(GetSettingValue("renumber.trackChangesMode", "tracked"))) = "silent" Then
        doc.TrackRevisions = False
    End If

    sorted = SortOps(ops)
    For i = UBound(sorted) To LBound(sorted) Step -1
        op = sorted(i)
        Select Case op(OP_KIND)
            Case OPK_FLAG
                FlagSpan doc, starts, segs, op, CStr(op(OP_TEXT))
                flagCount = flagCount + 1
            Case Else
                If ReplaceSpan(doc, texts, starts, segs, op) Then
                    If op(OP_KIND) = OPK_RELABEL Then
                        labelsChanged = labelsChanged + 1
                    Else
                        refsChanged = refsChanged + 1
                    End If
                Else
                    FlagSpan doc, starts, segs, op, _
                        CmtManualCheck(Mid$(texts(op(OP_PARA)), op(OP_START), op(OP_LEN)), CStr(op(OP_TEXT)))
                    flagCount = flagCount + 1
                End If
        End Select
    Next i

    RestoreState doc, viewState, oldTrack, undoOpen
    MsgBox MsgRenumberDone(claimCount, labelsChanged, refsChanged, flagCount, _
        CStr(GetSettingValue("check.commentAuthor", "PatentCheck"))), vbInformation, AppTitle()
    Exit Sub

Fail:
    errDesc = Err.Description
    If stateSaved Then RestoreState doc, viewState, oldTrack, undoOpen
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Private Sub RestoreState(ByVal doc As Document, ByVal viewState As Variant, _
        ByVal oldTrack As Boolean, ByRef undoOpen As Boolean)
    On Error Resume Next
    doc.TrackRevisions = oldTrack
    RestoreRevisionView doc, viewState
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    undoOpen = False
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- reading

' Reads every paragraph of the main text. For each paragraph i:
'   texts(i)  visible text (tracked deletions removed, no paragraph mark)
'   starts(i) document position of the paragraph start
'   segs(i)   Empty if visible text maps 1:1 onto the document, otherwise an
'             array of segments Array(visibleStart, docStart, length)
Private Sub ReadVisibleParagraphs(ByVal doc As Document, ByRef texts As Variant, _
        ByRef starts As Variant, ByRef segs As Variant)
    Dim delStart() As Long, delEnd() As Long, delCount As Long
    Dim t() As String, st() As Long, sg() As Variant
    Dim para As Paragraph, rng As Range, raw As String
    Dim i As Long, rs As Long, re As Long, ptr As Long

    CollectDeletions doc, delStart, delEnd, delCount
    ReDim t(0 To doc.Paragraphs.Count - 1)
    ReDim st(0 To doc.Paragraphs.Count - 1)
    ReDim sg(0 To doc.Paragraphs.Count - 1)

    For Each para In doc.Paragraphs
        Set rng = para.Range
        raw = rng.Text
        If Right$(raw, 1) = vbCr Then raw = Left$(raw, Len(raw) - 1)
        rs = rng.Start
        re = rs + Len(raw)
        st(i) = rs

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

' ---------------------------------------------------------------- writing

' Maps a visible span of paragraph p to document positions [a, b).
' Fails if the span is interrupted by tracked deletions.
Private Function MapSpan(ByVal starts As Variant, ByVal segs As Variant, ByVal p As Long, _
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

' Replaces one label; verifies the document still holds the expected text first.
Private Function ReplaceSpan(ByVal doc As Document, ByVal texts As Variant, ByVal starts As Variant, _
        ByVal segs As Variant, ByVal op As Variant) As Boolean
    Dim a As Long, b As Long, rng As Range
    If Not MapSpan(starts, segs, op(OP_PARA), op(OP_START), op(OP_LEN), a, b) Then Exit Function
    Set rng = doc.Range(a, b)
    If rng.Text <> Mid$(texts(op(OP_PARA)), op(OP_START), op(OP_LEN)) Then Exit Function
    rng.Text = op(OP_TEXT)
    ReplaceSpan = True
End Function

Private Sub FlagSpan(ByVal doc As Document, ByVal starts As Variant, ByVal segs As Variant, _
        ByVal op As Variant, ByVal commentText As String)
    Dim a As Long, b As Long
    If Not MapSpan(starts, segs, op(OP_PARA), op(OP_START), op(OP_LEN), a, b) Then
        a = starts(op(OP_PARA))
        b = a
    End If
    AddToolComment doc, a, b, commentText
End Sub
