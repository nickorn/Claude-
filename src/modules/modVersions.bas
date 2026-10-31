Attribute VB_Name = "modVersions"
Option Explicit

' Roadmap 4: clean version / marked-up version + footer status.
'   Reinschrift erzeugen       new document, all changes accepted, all comments removed
'   Markierte Fassung erzeugen new document with the Track Changes mark-up; tool marks
'                              removed, other comments per versions.markedUpKeepComments
'   Fusszeile/Status           sets the status in the current document's footer
' The original document is never changed. The copy is made from the saved
' file (Documents.Add with the file as template), so headers, footers, styles,
' tracked changes and comments come along; unsaved changes are saved first
' after asking.

Public Const STATUS_MARKED_UP As String = "markedUp"
Public Const STATUS_CLEAN As String = "clean"
Public Const STATUS_APPROVED As String = "inventorApproved"

Public Sub CreateCleanVersion()
    CreateVersion STATUS_CLEAN
End Sub

Public Sub CreateMarkedUpVersion()
    CreateVersion STATUS_MARKED_UP
End Sub

' statusKey: markedUp / clean / inventorApproved / "" (remove)
Public Sub SetFooterStatus(ByVal statusKey As String)
    Dim doc As Document, errDesc As String, n As Long, status As String
    Set doc = BeginDocWork(UndoTitleFooter())
    If doc Is Nothing Then Exit Sub
    On Error GoTo Fail
    SetTrackMode "silent"
    status = StatusText(statusKey, DocumentLanguage(doc))
    n = WriteFooterStatus(doc, status)
    EndDocWork
    MsgBox MsgFooterSet(status, n), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    EndDocWork
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Private Sub CreateVersion(ByVal statusKey As String)
    Dim doc As Document, copyDoc As Document, lang As String, errDesc As String, hadRevisions As Boolean
    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Sub
    End If
    Set doc = ActiveDocument
    ReloadSettings
    If Not EnsureSaved(doc) Then Exit Sub

    On Error GoTo Fail
    Application.ScreenUpdating = False
    lang = DocumentLanguage(doc)
    hadRevisions = (doc.Revisions.Count > 0)
    Set copyDoc = Documents.Add(Template:=doc.FullName, NewTemplate:=False)
    copyDoc.TrackRevisions = False

    RemoveToolMarks copyDoc
    If statusKey = STATUS_CLEAN Then
        copyDoc.AcceptAllRevisions
        If copyDoc.Comments.Count > 0 Then copyDoc.DeleteAllComments
    Else
        If Not CBool(GetSettingValue("versions.markedUpKeepComments", False)) Then
            If copyDoc.Comments.Count > 0 Then copyDoc.DeleteAllComments
        End If
        copyDoc.TrackRevisions = doc.TrackRevisions
    End If
    WriteFooterStatus copyDoc, StatusText(statusKey, lang)
    Application.ScreenUpdating = True
    copyDoc.Activate

    If statusKey = STATUS_MARKED_UP And hadRevisions And copyDoc.Revisions.Count = 0 Then
        MsgBox MsgVersionLostRevisions(), vbExclamation, AppTitle()
    ElseIf statusKey = STATUS_MARKED_UP And Not hadRevisions Then
        MsgBox MsgVersionNoRevisions(), vbInformation, AppTitle()
    Else
        MsgBox MsgVersionCreated(StatusText(statusKey, lang)), vbInformation, AppTitle()
    End If
    Exit Sub
Fail:
    errDesc = Err.Description
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' The copy is made from the file, so the file must be up to date.
Private Function EnsureSaved(ByVal doc As Document) As Boolean
    If Len(doc.Path) = 0 Then
        MsgBox MsgVersionSaveFirst(), vbExclamation, AppTitle()
        Exit Function
    End If
    If Not doc.Saved Then
        If MsgBox(MsgVersionAskSave(doc.Name), vbQuestion + vbYesNo, AppTitle()) <> vbYes Then Exit Function
        doc.Save
    End If
    EnsureSaved = True
End Function

' Claim language of the document (settings language.claimLanguageDetection).
Public Function DocumentLanguage(ByVal doc As Document) As String
    Dim texts As Variant, starts As Variant, segs As Variant, indents As Variant, setting As String
    setting = LCase$(CStr(GetSettingValue("language.claimLanguageDetection", "auto")))
    If setting = "en" Or setting = "de" Then
        DocumentLanguage = setting
        Exit Function
    End If
    ReadVisibleParagraphs doc, texts, starts, segs, indents
    BuildClaimModel texts, indents, "auto"
    If gClaimCount > 0 Then DocumentLanguage = gLanguage Else DocumentLanguage = "de"
End Function

Public Function StatusText(ByVal statusKey As String, ByVal lang As String) As String
    If Len(statusKey) = 0 Then Exit Function
    StatusText = CStr(GetSettingValue("footer.status." & lang & "." & statusKey, ""))
End Function

Private Function KnownStatusTexts() As String
    Dim langs As Variant, keys As Variant, i As Long, j As Long, s As String
    langs = Array("de", "en")
    keys = Array(STATUS_MARKED_UP, STATUS_CLEAN, STATUS_APPROVED)
    s = "|"
    For i = LBound(langs) To UBound(langs)
        For j = LBound(keys) To UBound(keys)
            s = s & StatusText(CStr(keys(j)), CStr(langs(i))) & "|"
        Next j
    Next i
    KnownStatusTexts = s
End Function

' Writes the status into every distinct footer of the document. Returns the
' number of footers changed.
Public Function WriteFooterStatus(ByVal doc As Document, ByVal status As String) As Long
    Dim sec As Section, kinds As Variant, i As Long, ft As HeaderFooter, known As String
    known = KnownStatusTexts()
    kinds = Array(wdHeaderFooterPrimary, wdHeaderFooterFirstPage, wdHeaderFooterEvenPages)
    For Each sec In doc.Sections
        For i = LBound(kinds) To UBound(kinds)
            If FooterInUse(sec, CLng(kinds(i))) Then
                Set ft = sec.Footers(CLng(kinds(i)))
                If sec.Index = 1 Or Not ft.LinkToPrevious Then
                    If WriteStatusIntoFooter(sec, ft, status, known) Then WriteFooterStatus = WriteFooterStatus + 1
                End If
            End If
        Next i
    Next sec
End Function

Private Function FooterInUse(ByVal sec As Section, ByVal kind As Long) As Boolean
    Select Case kind
        Case wdHeaderFooterPrimary: FooterInUse = True
        Case wdHeaderFooterFirstPage: FooterInUse = sec.PageSetup.DifferentFirstPageHeaderFooter
        Case wdHeaderFooterEvenPages: FooterInUse = sec.PageSetup.OddAndEvenPagesHeaderFooter
    End Select
End Function

Private Function WriteStatusIntoFooter(ByVal sec As Section, ByVal ft As HeaderFooter, _
        ByVal status As String, ByVal known As String) As Boolean
    Dim para As Paragraph, t As String, plan As Variant, r As Range, cellRange As Range

    ' two-cell table layout: status = first cell, everything else untouched
    If ft.Range.Tables.Count > 0 Then
        If ft.Range.Tables(1).Range.Cells.Count >= 2 Then
            Set cellRange = ft.Range.Tables(1).Range.Cells(1).Range
            cellRange.MoveEnd wdCharacter, -1 ' without the end-of-cell mark
            If cellRange.Text <> status Then
                cellRange.Text = status
                WriteStatusIntoFooter = True
            End If
            Exit Function
        End If
    End If

    Set para = ft.Range.Paragraphs(1)
    t = para.Range.Text
    If Right$(t, 1) = vbCr Then t = Left$(t, Len(t) - 1)
    plan = PlanFooterLeft(t, status, known)

    Select Case plan(0)
        Case FA_REPLACE, FA_REPLACE_EMPTY
            Set r = para.Range.Duplicate
            r.SetRange r.Start + plan(1) - 1, r.Start + plan(1) - 1 + plan(2)
            If r.Text <> Mid$(t, plan(1), plan(2)) Then
                ' fields etc. make positions unreliable: do not risk the right part
                InsertStatusParagraph para, status
            Else
                r.Text = plan(3)
                If plan(0) = FA_REPLACE_EMPTY Then AddRightTabStop sec, para
            End If
            WriteStatusIntoFooter = True
        Case FA_INSERT_BEFORE
            InsertStatusParagraph para, status
            WriteStatusIntoFooter = True
    End Select
End Function

' New first footer paragraph with the status; the existing paragraph (e.g. the
' application number) stays as it is. Range.InsertParagraphBefore expands the
' range to include the new paragraph, which is then its first paragraph.
Private Sub InsertStatusParagraph(ByVal para As Paragraph, ByVal status As String)
    Dim r As Range, np As Range
    If Len(status) = 0 Then Exit Sub
    Set r = para.Range.Duplicate
    r.InsertParagraphBefore
    Set np = r.Paragraphs(1).Range
    np.MoveEnd wdCharacter, -1
    np.Text = status
    np.ParagraphFormat.Alignment = wdAlignParagraphLeft
End Sub

Private Sub AddRightTabStop(ByVal sec As Section, ByVal para As Paragraph)
    Dim w As Single
    With sec.PageSetup
        w = .PageWidth - .LeftMargin - .RightMargin
    End With
    para.TabStops.Add Position:=w, Alignment:=wdAlignTabRight
End Sub
