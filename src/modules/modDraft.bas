Attribute VB_Name = "modDraft"
Option Explicit

' Roadmap 6 drafting helpers (Word layer):
'   Anspruchsbaum                 dependency tree as a new document
'   Abhaengigen Anspruch anlegen  skeleton after the claim at the cursor
'   Ansprueche zusammenfuehren    merge the dependent claim at the cursor into
'                                 its parent, redirect references, renumber
' Text changes follow the document's Track Changes state (renumber.trackChangesMode).

Public Sub ShowClaimTree()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim lines As Collection, ln As Variant, body As String, rep As Document, para As Paragraph
    Dim depths() As Long, i As Long, indep As Long, c As Long, errDesc As String

    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Sub
    End If
    Set doc = ActiveDocument
    On Error GoTo Fail
    ReloadSettings
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If Not BuildClaimModel(texts, indents, CStr(GetSettingValue("language.claimLanguageDetection", "auto"))) Then
        MsgBox MsgNoClaims(), vbExclamation, AppTitle()
        Exit Sub
    End If
    For c = 1 To gClaimCount
        If Not IsDependentClaim(c) Then indep = indep + 1
    Next c

    Application.ScreenUpdating = False
    Set lines = ClaimTreeLines()
    ReDim depths(1 To lines.Count + 3)
    body = TreeTitle(doc.Name) & vbCr & TreeSummary(gClaimCount, indep) & vbCr
    i = 3
    For Each ln In lines
        i = i + 1
        depths(i) = ln(TL_DEPTH)
        body = body & vbCr & ln(TL_LABEL) & "." & vbTab & ln(TL_TEXT)
        If Len(ln(TL_NOTE)) > 0 Then body = body & "  [" & ln(TL_NOTE) & "]"
    Next ln

    Set rep = Documents.Add
    rep.Content.Text = body
    rep.Content.Font.Name = CStr(GetSettingValue("format.fontName", "Arial"))
    rep.Content.Font.Size = 10
    rep.Content.ParagraphFormat.SpaceAfter = 0
    i = 0
    For Each para In rep.Paragraphs
        i = i + 1
        If i > 3 And i <= UBound(depths) Then
            para.LeftIndent = CentimetersToPoints(0.8 * depths(i)) + CentimetersToPoints(1)
            para.FirstLineIndent = -CentimetersToPoints(1)
            para.TabStops.Add Position:=para.LeftIndent
        End If
    Next para
    rep.Paragraphs(1).Range.Font.Bold = True
    rep.Paragraphs(1).Range.Font.Size = 12
    rep.Saved = True
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    errDesc = Err.Description
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Public Sub AddDependentClaim()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim c As Long, label As String, newText As String, lastP As Range, pos As Long, errDesc As String
    Dim claimFormat As ParagraphFormat

    Set doc = BeginDocWork(UndoTitleDraft(1))
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If Not BuildClaimModel(texts, indents, CStr(GetSettingValue("language.claimLanguageDetection", "auto"))) Then
        EndDocWork
        MsgBox MsgNoClaims(), vbExclamation, AppTitle()
        Exit Sub
    End If
    c = ClaimAtSelection(doc, starts)
    If c = 0 Then
        EndDocWork
        MsgBox MsgNoClaimAtCursor(), vbExclamation, AppTitle()
        Exit Sub
    End If
    SetTrackMode CStr(GetSettingValue("renumber.trackChangesMode", "tracked"))

    label = NextSupplementaryLabel(gClaimLabel(c))
    newText = label & "." & vbTab & DependentSkeleton(c)
    ' after the claim's last paragraph: [last]<new empty><new claim><old mark>
    Set lastP = doc.Paragraphs(gClaimLastPara(c) + 1).Range
    Set claimFormat = doc.Paragraphs(gClaimPara(c) + 1).Format.Duplicate
    pos = lastP.End - 1
    doc.Range(pos, pos).InsertAfter vbCr & vbCr & newText
    doc.Range(pos + 1, pos + 2).Paragraphs(1).Format = claimFormat
    doc.Range(pos + 2, pos + 2 + Len(newText)).Paragraphs(1).Format = claimFormat
    EndDocWork
    Selection.SetRange pos + 2 + Len(newText), pos + 2 + Len(newText)
    MsgBox MsgSkeletonInserted(label), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Public Sub MergeClaimIntoParent()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim d As Long, plan As Variant, parent As Long, dLabel As String, pLabel As String, lang As String
    Dim ops As New Collection, lastText As String, failed As Long, result As Variant, errDesc As String
    Dim c As Long, r As Variant

    Set doc = BeginDocWork(UndoTitleDraft(2))
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    lang = CStr(GetSettingValue("language.claimLanguageDetection", "auto"))
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If Not BuildClaimModel(texts, indents, lang) Then
        EndDocWork
        MsgBox MsgNoClaims(), vbExclamation, AppTitle()
        Exit Sub
    End If
    LoadCheckOptions
    d = ClaimAtSelection(doc, starts)
    plan = PlanMerge(d)
    If Not plan(0) Then
        EndDocWork
        MsgBox plan(1), vbExclamation, AppTitle()
        Exit Sub
    End If
    parent = plan(2)
    dLabel = gClaimLabel(d)
    pLabel = gClaimLabel(parent)
    Application.ScreenUpdating = True
    If MsgBox(MsgMergeConfirm(dLabel, pLabel), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then
        EndDocWork
        Exit Sub
    End If
    Application.ScreenUpdating = False
    SetTrackMode CStr(GetSettingValue("renumber.trackChangesMode", "tracked"))

    ' 1. remove claim d (it comes after its parent, so the parent's positions stay valid)
    If d < gClaimCount Then
        doc.Range(starts(gClaimPara(d)), starts(gClaimPara(d + 1))).Delete
    Else
        doc.Range(starts(gClaimLastPara(d - 1) + 1) - 1, doc.Content.End - 1).Delete
    End If

    ' 2. append d's features to the parent
    lastText = RTrimText(texts(gClaimLastPara(parent)))
    ops.Add Array(OPK_REFERENCE, gClaimLastPara(parent), Len(lastText), 1, plan(3))
    If plan(4) >= 0 Then
        ops.Add Array(OPK_REFERENCE, plan(4), Len(RTrimText(texts(plan(4)))) - plan(5) + 1, plan(5), "")
    End If
    ApplyTextOps doc, texts, starts, segs, ops, failed

    ' 3. references to d now point to the parent
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    BuildClaimModel texts, indents, lang
    Set ops = New Collection
    For c = 1 To gClaimCount
        For Each r In gClaimRefs(c)
            If r(R_LABEL) = dLabel And r(R_TARGET) = 0 Then
                ops.Add Array(OPK_REFERENCE, r(R_PARA), r(R_START), r(R_LEN), pLabel)
            End If
        Next r
    Next c
    ApplyTextOps doc, texts, starts, segs, ops, failed

    ' 4. renumber the rest
    result = RenumberDocument(doc)
    EndDocWork
    MsgBox MsgMergeDone(dLabel, pLabel, gLanguage) & vbCr & vbCr & result(1), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' Claim containing the cursor (0 = before the first claim, or the cursor is
' not in the main text, e.g. in a footer or comment).
Public Function ClaimAtSelection(ByVal doc As Document, ByVal starts As Variant) As Long
    Dim p As Long, sel As Long, hit As Long
    If Selection.StoryType <> wdMainTextStory Then Exit Function
    sel = Selection.Start
    hit = -1
    For p = LBound(starts) To UBound(starts)
        If starts(p) > sel Then Exit For
        hit = p
    Next p
    If hit >= 0 Then ClaimAtSelection = gClaimOfPara(hit)
End Function
