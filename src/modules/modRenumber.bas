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
' Track Changes (settings: renumber.trackChangesMode, applied via SetTrackMode):
'   "tracked" - changes follow the document's Track Changes state (recorded if on)
'   "silent"  - Track Changes is switched off while renumbering

Public Sub RenumberClaims()
    Dim doc As Document, result As Variant, errDesc As String
    Set doc = BeginDocWork(UndoTitleRenumber())
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    result = RenumberDocument(doc)
    EndDocWork
    MsgBox result(1), result(2), AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' Renumbers the document; the caller manages the work session (BeginDocWork),
' so other commands (e.g. claim merge) can renumber inside their own undo step.
' Returns Array(ok, message, msgBoxStyle).
Public Function RenumberDocument(ByVal doc As Document) As Variant
    Dim texts As Variant, starts As Variant, segs As Variant
    Dim ops As New Collection, sorted As Variant, op As Variant
    Dim errMsg As String, claimCount As Long, i As Long
    Dim labelsChanged As Long, refsChanged As Long, flagCount As Long

    ReadVisibleParagraphs doc, texts, starts, segs
    If Not PlanRenumber(texts, ops, errMsg, claimCount) Then
        RenumberDocument = Array(False, errMsg, vbExclamation)
        Exit Function
    End If
    If ops.Count = 0 Then
        RenumberDocument = Array(True, MsgAlreadyConsecutive(claimCount), vbInformation)
        Exit Function
    End If

    SetTrackMode CStr(GetSettingValue("renumber.trackChangesMode", "tracked"))
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
    EndCommentBatch
    RenumberDocument = Array(True, MsgRenumberDone(claimCount, labelsChanged, refsChanged, flagCount, ToolAuthor()), vbInformation)
End Function

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
