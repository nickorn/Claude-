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
    BeginCommentBatch
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
    EndCommentBatch
    doc.TrackRevisions = oldTrack
    RestoreRevisionView doc, viewState
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    undoOpen = False
    Application.ScreenUpdating = True
End Sub

' ---------------------------------------------------------------- writing

' Replaces one label; verifies the document still holds the expected text first.
Private Function ReplaceSpan(ByVal doc As Document, ByVal texts As Variant, ByVal starts As Variant, _
        ByVal segs As Variant, ByVal op As Variant) As Boolean
    ReplaceSpan = ApplyReplaceOp(doc, texts, starts, segs, op)
End Function

Private Sub FlagSpan(ByVal doc As Document, ByVal starts As Variant, ByVal segs As Variant, _
        ByVal op As Variant, ByVal commentText As String)
    Dim a As Long, b As Long
    If Not MapSpan(starts, segs, op(OP_PARA), op(OP_START), op(OP_LEN), a, b) Then
        a = starts(op(OP_PARA))
        b = a
    End If
    AddToolComment doc, a, b, "[" & CheckDisplayName("renumber") & "] " & commentText
End Sub
