Attribute VB_Name = "modConvert"
Option Explicit

' Roadmap 7 drafting helpers (Word layer):
'   Verfahren <-> Vorrichtung   counterpart of the independent claim at the cursor,
'                               appended as new last claim (draft to review)
'   Ansprueche -> Beschreibung  new document with one paragraph per claim and a
'                               selectable intro phrase (settings description.introPhrases)
'   Bezugszeichen entfernen     removes all sign groups in the claims
'   Bezugszeichen einfuegen     adds missing signs at the first mention per claim
'                               (map from this document or from another document)

Public Sub ConvertMethodApparatus()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim c As Long, res As Variant, paras As Variant, label As String, insertText As String, i As Long
    Dim lastP As Range, pos As Long, formats() As ParagraphFormat, errDesc As String, allInt As Boolean, np As Paragraph

    Set doc = BeginDocWork(UndoTitleConvert(1))
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
    res = ConvertClaimCategory(c)
    If Not res(0) Then
        EndDocWork
        MsgBox res(1), vbExclamation, AppTitle()
        Exit Sub
    End If
    paras = res(2)
    SetTrackMode CStr(GetSettingValue("renumber.trackChangesMode", "tracked"))

    ' label: next integer if the claims are numbered 1..n, else a supplementary label
    allInt = True
    For i = 1 To gClaimCount
        If gClaimLabel(i) <> CStr(i) Then allInt = False
    Next i
    If allInt Then label = CStr(gClaimCount + 1) Else label = NextSupplementaryLabel(gClaimLabel(gClaimCount))

    ' paragraph formats of the source claim (indents of the features)
    ReDim formats(0 To UBound(paras))
    For i = 0 To UBound(paras)
        Set formats(i) = doc.Paragraphs(gClaimPara(c) + i + 1).Format.Duplicate
    Next i

    insertText = vbCr & vbCr & label & "." & vbTab & paras(0)
    For i = 1 To UBound(paras)
        insertText = insertText & vbCr & paras(i)
    Next i
    Set lastP = doc.Paragraphs(gClaimLastPara(gClaimCount) + 1).Range
    pos = lastP.End - 1
    doc.Range(pos, pos).InsertAfter insertText

    ' new paragraphs: empty separator (claim format), then the claim paragraphs
    doc.Range(pos + 1, pos + 1).Paragraphs(1).Format = formats(0)
    Set np = doc.Range(pos + 2, pos + 2).Paragraphs(1)
    For i = 0 To UBound(paras)
        If np Is Nothing Then Exit For
        np.Format = formats(i)
        Set np = np.Next
    Next i
    EndDocWork
    MsgBox MsgConvertDone(label), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Public Sub ClaimsToDescription()
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim intro As String, body As String, c As Long, rep As Document, errDesc As String

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
    intro = ChooseIntroPhrase(gLanguage)
    If intro = Chr$(0) Then Exit Sub

    For c = 1 To gClaimCount
        If c > 1 Then body = body & vbCr & vbCr
        body = body & ClaimToDescription(c, intro)
    Next c
    Application.ScreenUpdating = False
    Set rep = Documents.Add
    rep.Content.Text = body
    With rep.Content
        .Font.Name = CStr(GetSettingValue("format.fontName", "Arial"))
        .Font.Size = CSng(GetSettingValue("format.fontSize", 11))
        .ParagraphFormat.LineSpacingRule = wdLineSpace1pt5
        .ParagraphFormat.SpaceBefore = 0
        .ParagraphFormat.SpaceAfter = 0
    End With
    Application.ScreenUpdating = True
    MsgBox MsgDescriptionDone(gClaimCount), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' Lets the user pick one of description.introPhrases.<lang>; Chr(0) = cancelled.
Private Function ChooseIntroPhrase(ByVal lang As String) As String
    Dim root As Object, list As Collection, v As Variant, prompt As String, i As Long, answer As String
    Set root = JsonParse(DecodeJsonUnicode(DefaultSettingsJson()))
    On Error Resume Next
    Set root = JsonParse(ReadUtf8File(SettingsPath()))
    On Error GoTo 0
    Set list = New Collection
    If root.Exists("description") Then
        If root.Item("description").Exists("introPhrases") Then
            If root.Item("description").Item("introPhrases").Exists(lang) Then
                For Each v In root.Item("description").Item("introPhrases").Item(lang)
                    list.Add CStr(v)
                Next v
            End If
        End If
    End If
    If list.Count = 0 Then
        If lang = "de" Then list.Add Txt("Gem{ae}{ss} einer Ausf{ue}hrungsform ") Else list.Add "According to an embodiment, "
    End If
    For i = 1 To list.Count
        prompt = prompt & i & " = " & list(i) & ChrW(8230) & vbCr
    Next i
    answer = InputBox(MsgChooseIntro() & vbCr & vbCr & prompt, AppTitle(), "1")
    If Len(answer) = 0 Then
        ChooseIntroPhrase = Chr$(0)
    ElseIf IsNumeric(answer) Then
        i = CLng(answer)
        If i < 1 Or i > list.Count Then i = 1
        ChooseIntroPhrase = list(i)
    Else
        ChooseIntroPhrase = answer & IIf(Right$(answer, 1) = " ", "", " ")
    End If
End Function

Public Sub StripReferenceSigns()
    RunSignCommand False
End Sub

Public Sub InsertReferenceSigns()
    RunSignCommand True
End Sub

Private Sub RunSignCommand(ByVal insertMode As Boolean)
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim ops As New Collection, map As Collection, applied As Long, failed As Long, errDesc As String, lang As String
    Dim orig As Document

    Set doc = BeginDocWork(UndoTitleConvert(IIf(insertMode, 3, 2)))
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    lang = CStr(GetSettingValue("language.claimLanguageDetection", "auto"))
    LoadCheckOptions
    If insertMode Then
        ' map from this document; if it has no signs, from another document
        ReadVisibleParagraphs doc, texts, starts, segs, indents
        If Not BuildClaimModel(texts, indents, lang) Then GoTo NoClaims
        Set map = CollectSignMap()
        If map.Count = 0 Then
            EndDocWork
            Set orig = doc
            Set map = SignMapFromOtherDocument(lang)
            If map Is Nothing Then Exit Sub
            orig.Activate
            Set doc = BeginDocWork(UndoTitleConvert(3))
            If doc Is Nothing Then Exit Sub
        End If
    End If
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If Not BuildClaimModel(texts, indents, lang) Then GoTo NoClaims
    LoadCheckOptions
    If insertMode Then PlanInsertSigns ops, map Else PlanStripSigns ops
    SetTrackMode CStr(GetSettingValue("format.textTrackChanges", "tracked"))
    applied = ApplyTextOps(doc, texts, starts, segs, ops, failed)
    EndDocWork
    MsgBox MsgSignsDone(insertMode, applied, failed), vbInformation, AppTitle()
    Exit Sub
NoClaims:
    EndDocWork
    MsgBox MsgNoClaims(), vbExclamation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' Opens another document read-only (e.g. the EP version with signs) and
' collects its term -> sign map. Nothing if cancelled or no signs found.
Private Function SignMapFromOtherDocument(ByVal lang As String) As Collection
    Dim path As String, src As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim map As Collection, d As Document
    If MsgBox(MsgSignsAskSource(), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Function
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = Txt("Dokument mit Bezugszeichen w{ae}hlen")
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Function
        path = .SelectedItems(1)
    End With
    ' Documents.Open would return an already open document - and closing it
    ' afterwards would close the user's work. Only read files that are not open.
    For Each d In Documents
        If LCase$(d.FullName) = LCase$(path) Then
            MsgBox MsgSignsSourceOpen(d.Name), vbExclamation, AppTitle()
            Exit Function
        End If
    Next d
    Set src = Documents.Open(FileName:=path, ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
    ReadVisibleParagraphs src, texts, starts, segs, indents
    If BuildClaimModel(texts, indents, lang) Then Set map = CollectSignMap()
    src.Close SaveChanges:=wdDoNotSaveChanges
    If map Is Nothing Then
        MsgBox MsgSignsNoneFound(), vbExclamation, AppTitle()
    ElseIf map.Count = 0 Then
        MsgBox MsgSignsNoneFound(), vbExclamation, AppTitle()
        Set map = Nothing
    End If
    Set SignMapFromOtherDocument = map
End Function

