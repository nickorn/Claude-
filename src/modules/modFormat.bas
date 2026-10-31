Attribute VB_Name = "modFormat"
Option Explicit

' MVP-3 formatting fixer.
'   Formatierung anwenden : font, size, line spacing, space before/after,
'                           sub-feature indents, exactly one empty paragraph
'                           between claims (layout; settings format.*)
'   Doppelte Leerzeichen  : modFormatText.PlanSpaceFixes
'   Zeichensetzung        : modFormatText.PlanPunctuationFixes
'   Bezugszeichen-Format  : modFormatText.PlanSignFormatFixes
' Layout changes follow format.layoutTrackChanges (default "silent"), text
' fixes follow format.textTrackChanges (default "tracked" = document state).

Private Const FIX_SPACES As Long = 1
Private Const FIX_PUNCTUATION As Long = 2
Private Const FIX_SIGNS As Long = 3

Public Sub FixSpaces()
    RunTextFix FIX_SPACES
End Sub

Public Sub FixPunctuation()
    RunTextFix FIX_PUNCTUATION
End Sub

Public Sub FixSignFormat()
    RunTextFix FIX_SIGNS
End Sub

Private Sub RunTextFix(ByVal kind As Long)
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim ops As New Collection, applied As Long, failed As Long, errDesc As String, hasClaims As Boolean

    Set doc = BeginDocWork(UndoTitleFormat(kind))
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail

    ReadVisibleParagraphs doc, texts, starts, segs, indents
    hasClaims = BuildClaimModel(texts, indents, CStr(GetSettingValue("language.claimLanguageDetection", "auto")))
    If Not hasClaims And kind <> FIX_SPACES Then
        EndDocWork
        MsgBox MsgNoClaims(), vbExclamation, AppTitle()
        Exit Sub
    End If
    LoadCheckOptions

    Select Case kind
        Case FIX_SPACES: PlanSpaceFixes ops
        Case FIX_PUNCTUATION: PlanPunctuationFixes ops
        Case FIX_SIGNS: PlanSignFormatFixes ops
    End Select

    SetTrackMode CStr(GetSettingValue("format.textTrackChanges", "tracked"))
    applied = ApplyTextOps(doc, texts, starts, segs, ops, failed)
    EndDocWork
    MsgBox MsgFormatTextDone(kind, applied, failed), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' ---------------------------------------------------------------- layout

Public Sub ApplyFormatting()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim errDesc As String, added As Long, removed As Long, indented As Long, rng As Range
    Dim lang As String

    Set doc = BeginDocWork(UndoTitleFormat(0))
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    SetTrackMode CStr(GetSettingValue("format.layoutTrackChanges", "silent"))
    lang = CStr(GetSettingValue("language.claimLanguageDetection", "auto"))

    ' 1. empty paragraphs between claims (structure first, positions from the end)
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If BuildClaimModel(texts, indents, lang) And CBool(GetSettingValue("format.fixEmptyParagraphs", True)) Then
        FixEmptyParagraphs doc, starts, segs, added, removed
    End If

    ' 2. font and paragraph spacing
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    BuildClaimModel texts, indents, lang
    If CBool(GetSettingValue("format.applyToWholeDocument", True)) Or gClaimCount = 0 Then
        Set rng = doc.Content
    Else
        Set rng = doc.Range(starts(gClaimPara(1)), doc.Content.End)
    End If
    ApplyCharacterAndParagraphFormat rng

    ' 3. sub-feature indents to whole multiples of the configured depth
    indented = NormaliseIndents(doc, texts)

    EndDocWork
    MsgBox MsgFormatLayoutDone(added, removed, indented), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Private Sub ApplyCharacterAndParagraphFormat(ByVal rng As Range)
    Dim factor As Single
    rng.Font.Name = CStr(GetSettingValue("format.fontName", "Arial"))
    rng.Font.Size = CSng(GetSettingValue("format.fontSize", 11))
    factor = CSng(GetSettingValue("format.lineSpacing", 1.5))
    With rng.ParagraphFormat
        If Abs(factor - 1.5) < 0.01 Then
            .LineSpacingRule = wdLineSpace1pt5
        ElseIf Abs(factor - 1) < 0.01 Then
            .LineSpacingRule = wdLineSpaceSingle
        ElseIf Abs(factor - 2) < 0.01 Then
            .LineSpacingRule = wdLineSpaceDouble
        Else
            .LineSpacingRule = wdLineSpaceMultiple
            .LineSpacing = LinesToPoints(factor)
        End If
        .SpaceBeforeAuto = False
        .SpaceAfterAuto = False
        .SpaceBefore = CSng(GetSettingValue("format.spaceBefore", 0))
        .SpaceAfter = CSng(GetSettingValue("format.spaceAfter", 0))
    End With
End Sub

' Exactly one empty paragraph between two claims. Works from the last claim
' backwards so earlier positions stay valid. Paragraphs with tracked
' deletions are left alone.
Private Sub FixEmptyParagraphs(ByVal doc As Document, ByVal starts As Variant, ByVal segs As Variant, _
        ByRef added As Long, ByRef removed As Long)
    Dim c As Long, gap As Long, p As Long
    For c = gClaimCount To 2 Step -1
        gap = gClaimPara(c) - gClaimLastPara(c - 1) - 1
        If gap = 0 Then
            doc.Range(starts(gClaimPara(c)), starts(gClaimPara(c))).InsertParagraphBefore
            added = added + 1
        ElseIf gap > 1 Then
            ' keep the first empty paragraph, delete the others (last first)
            For p = gClaimPara(c) - 1 To gClaimLastPara(c - 1) + 2 Step -1
                If IsEmpty(segs(p)) Then
                    doc.Range(starts(p), starts(p + 1)).Delete
                    removed = removed + 1
                End If
            Next p
        End If
    Next c
End Sub

' Indented paragraphs inside claims get a left indent of a whole multiple of
' format.subFeatureIndentCm (nearest level, at least 1) and no first-line indent.
Private Function NormaliseIndents(ByVal doc As Document, ByVal texts As Variant) As Long
    Dim para As Paragraph, i As Long, unit As Single, level As Long, target As Single
    unit = CentimetersToPoints(CSng(GetSettingValue("format.subFeatureIndentCm", 1.25)))
    If unit <= 0 Then Exit Function
    For Each para In doc.Paragraphs
        If gClaimOfPara(i) > 0 And Len(Trim$(texts(i))) > 0 Then
            If gClaimPara(gClaimOfPara(i)) <> i And para.LeftIndent > 0.5 Then
                level = CLng(para.LeftIndent / unit)
                If level < 1 Then level = 1
                target = level * unit
                If Abs(para.LeftIndent - target) > 0.5 Or Abs(para.FirstLineIndent) > 0.5 Then
                    para.LeftIndent = target
                    para.FirstLineIndent = 0
                    NormaliseIndents = NormaliseIndents + 1
                End If
            End If
        End If
        i = i + 1
    Next para
End Function
