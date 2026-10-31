Attribute VB_Name = "modCheck"
Option Explicit

' MVP-2: final claim check.
' Text checks run on the claim model (modCheckStructure, modCheckTerms), the
' formatting check on Word paragraphs (modCheckFormat). Findings are marked as
' comments of the tool author ("PatentCheck") with a highlight colour per check
' (settings: check.highlightColors), or listed in a report document.
' Highlights are set with Track Changes off, so no tracked formatting appears.

Private Const MODE_MARK As Long = 1
Private Const MODE_REPORT As Long = 2

Public Sub RunFullCheck()
    RunChecks "", MODE_MARK
End Sub

Public Sub RunSingleCheck(ByVal checkKey As String)
    RunChecks checkKey, MODE_MARK
End Sub

Public Sub RunCheckReport()
    RunChecks "", MODE_REPORT
End Sub

Public Sub RemoveCheckMarks()
    Dim doc As Document, oldTrack As Boolean, n As Long, errDesc As String
    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Sub
    End If
    Set doc = ActiveDocument
    ReloadSettings
    oldTrack = doc.TrackRevisions
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.UndoRecord.StartCustomRecord UndoTitleRemoveMarks()
    doc.TrackRevisions = False
    n = RemoveToolMarks(doc)
    doc.TrackRevisions = oldTrack
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox MsgMarksRemoved(n), vbInformation, AppTitle()
    Exit Sub
Fail:
    errDesc = Err.Description
    On Error Resume Next
    doc.TrackRevisions = oldTrack
    Application.UndoRecord.EndCustomRecord
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

' Loads the check options from settings.json into the pure check modules.
Public Sub LoadCheckOptions()
    SetDefaultCheckOptions
    gOptEnforceAnd = CBool(GetSettingValue("punctuation.enforceAndBeforeLastFeature", True))
    gOptFlagWherein = CBool(GetSettingValue("punctuation.flagWhereinAfterFullStop", True))
    gOptRefSignStyle = LCase$(CStr(GetSettingValue("referenceSigns.style", "parentheses")))
    gOptRefSignSeparator = CStr(GetSettingValue("referenceSigns.multipleSeparator", ", "))
    gOptNormaliseSemicolon = CBool(GetSettingValue("referenceSigns.normaliseSemicolon", True))
    gOptFlagEveryMultiDep = CBool(GetSettingValue("check.flagEveryMultipleDependency", False))
    Set gBlacklist = GetBlacklistTerms(gLanguage)
End Sub

Private Sub RunChecks(ByVal onlyKey As String, ByVal mode As Long)
    Dim doc As Document, texts As Variant, starts As Variant, segs As Variant, indents As Variant
    Dim findings As New Collection, keys As Variant, i As Long, key As String
    Dim viewState As Variant, oldTrack As Boolean, undoOpen As Boolean, stateSaved As Boolean
    Dim errDesc As String, makeReport As Boolean, marked As Boolean

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

    ReadVisibleParagraphs doc, texts, starts, segs, indents
    If Not BuildClaimModel(texts, indents, CStr(GetSettingValue("language.claimLanguageDetection", "auto"))) Then
        RestoreCheckState doc, viewState, oldTrack, undoOpen
        MsgBox MsgNoClaims(), vbExclamation, AppTitle()
        Exit Sub
    End If
    LoadCheckOptions

    keys = Split(CHECK_KEYS, "|")
    For i = LBound(keys) To UBound(keys)
        key = keys(i)
        If ShouldRun(key, onlyKey) Then
            If key = "formatting" Then
                CheckFormatting doc, texts, findings
            Else
                RunTextCheck key, findings
            End If
        End If
    Next i

    makeReport = (mode = MODE_REPORT)
    If mode = MODE_MARK And Len(onlyKey) = 0 Then makeReport = CBool(GetSettingValue("check.createReport", False))

    If mode = MODE_MARK Then
        Application.UndoRecord.StartCustomRecord UndoTitleCheck()
        undoOpen = True
        doc.TrackRevisions = False
        RemoveToolMarks doc
        MarkFindings doc, texts, starts, segs, findings
        marked = True
    End If
    RestoreCheckState doc, viewState, oldTrack, undoOpen

    If makeReport Then CreateReport doc, findings
    MsgBox MsgCheckSummary(SummaryLines(findings), findings.Count, marked), vbInformation, AppTitle()
    Exit Sub

Fail:
    errDesc = Err.Description
    If stateSaved Then RestoreCheckState doc, viewState, oldTrack, undoOpen
    Application.ScreenUpdating = True
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Private Function ShouldRun(ByVal key As String, ByVal onlyKey As String) As Boolean
    If Len(onlyKey) > 0 Then
        ShouldRun = (key = onlyKey)
    Else
        ShouldRun = CBool(GetSettingValue("check.enabled." & key, True))
    End If
End Function

Private Sub RestoreCheckState(ByVal doc As Document, ByVal viewState As Variant, _
        ByVal oldTrack As Boolean, ByRef undoOpen As Boolean)
    On Error Resume Next
    EndCommentBatch
    doc.TrackRevisions = oldTrack
    RestoreRevisionView doc, viewState
    If undoOpen Then Application.UndoRecord.EndCustomRecord
    undoOpen = False
    Application.ScreenUpdating = True
End Sub

' Marks findings from the end of the document backwards.
Private Sub MarkFindings(ByVal doc As Document, ByVal texts As Variant, ByVal starts As Variant, _
        ByVal segs As Variant, ByVal findings As Collection)
    Dim sorted As Variant, i As Long, f As Variant, p As Long, st As Long, ln As Long, a As Long, b As Long
    sorted = SortFindings(findings)
    BeginCommentBatch
    For i = UBound(sorted) To LBound(sorted) Step -1
        f = sorted(i)
        p = f(F_PARA)
        st = f(F_START)
        ln = f(F_LEN)
        If st > Len(texts(p)) Then st = Len(texts(p))
        If ln = 0 And Len(texts(p)) > 0 Then ln = 1
        If st < 1 Then
            a = starts(p)
            b = a
        ElseIf Not MapSpan(starts, segs, p, st, ln, a, b) Then
            a = starts(p)
            b = a
        End If
        If b > a Then doc.Range(a, b).HighlightColorIndex = _
            HighlightIndexFromName(CStr(GetSettingValue("check.highlightColors." & f(F_CHECK), "wdYellow")))
        AddToolComment doc, a, b, "[" & CheckDisplayName(CStr(f(F_CHECK))) & "] " & f(F_MSG)
    Next i
    EndCommentBatch
End Sub

Private Function SummaryLines(ByVal findings As Collection) As String
    Dim keys As Variant, i As Long, n As Long, s As String
    keys = Split(CHECK_KEYS, "|")
    For i = LBound(keys) To UBound(keys)
        n = CountFindings(findings, CStr(keys(i)))
        If n > 0 Then s = s & CheckDisplayName(CStr(keys(i))) & ": " & n & vbCr
    Next i
    SummaryLines = s
End Function

' ---------------------------------------------------------------- report

Private Sub CreateReport(ByVal doc As Document, ByVal findings As Collection)
    Dim rep As Document, sorted As Variant, i As Long, f As Variant, body As String, rng As Range, tbl As Table
    Set rep = Documents.Add
    If findings.Count = 0 Then
        rep.Content.Text = ReportTitle(doc.Name) & vbCr & vbCr & ReportNoFindings()
    Else
        sorted = SortFindings(findings)
        body = ReportHeader()
        For i = LBound(sorted) To UBound(sorted)
            f = sorted(i)
            body = body & vbCr & CellText(ClaimLabelOfPara(f(F_PARA))) & vbTab & _
                CellText(CheckDisplayName(CStr(f(F_CHECK)))) & vbTab & CellText(CStr(f(F_MSG))) & vbTab & _
                CellText(Excerpt(CLng(f(F_PARA)), CLng(f(F_START)), CLng(f(F_LEN))))
        Next i
        rep.Content.Text = ReportTitle(doc.Name) & vbCr & vbCr & body
        Set rng = rep.Range(rep.Paragraphs(3).Range.Start, rep.Content.End)
        Set tbl = rng.ConvertToTable(Separator:=wdSeparateByTabs, NumColumns:=4)
        tbl.Borders.Enable = True
        tbl.Rows(1).Range.Font.Bold = True
        tbl.Rows(1).HeadingFormat = True
    End If
    rep.Content.Font.Name = CStr(GetSettingValue("format.fontName", "Arial"))
    rep.Content.Font.Size = 10
    rep.Paragraphs(1).Range.Font.Size = 12
    rep.Paragraphs(1).Range.Font.Bold = True
    rep.Saved = True
End Sub

Private Function CellText(ByVal s As String) As String
    CellText = Replace(Replace(Replace(s, vbTab, " "), vbCr, " "), vbLf, " ")
End Function

Private Function Excerpt(ByVal p As Long, ByVal st As Long, ByVal ln As Long) As String
    Dim t As String, a As Long, e As Long
    t = gTexts(p)
    a = st - 30
    If a < 1 Then a = 1
    e = st + ln + 30
    If e > Len(t) + 1 Then e = Len(t) + 1
    Excerpt = Mid$(t, a, e - a)
    If a > 1 Then Excerpt = ChrW(8230) & Excerpt
    If e <= Len(t) Then Excerpt = Excerpt & ChrW(8230)
End Function
