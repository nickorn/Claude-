Attribute VB_Name = "modTextBlocks"
Option Explicit

' Roadmap 5: text blocks (Textbausteine).
' The ribbon's dynamic menu is rebuilt from the folder (settings
' textblocks.folder) every time it opens, so new files appear without
' restarting Word. One .docx per block (keeps formatting); subfolders are
' submenus. Placeholders [DATUM], [D1], ... are asked for on insert
' (textblocks.askForPlaceholders); unfilled ones are highlighted
' (textblocks.highlightUnfilledPlaceholders).

Public Function TextBlocksFolder() As String
    Dim f As String
    f = ExpandEnvironment(CStr(GetSettingValue("textblocks.folder", "%APPDATA%\PatentTools\textblocks")))
    If Right$(f, 1) = "\" Then f = Left$(f, Len(f) - 1)
    TextBlocksFolder = f
End Function

' Ribbon getContent: XML of the dynamic menu.
Public Function TextBlocksMenuContent() As String
    Dim base As String, listing As String, extra As String
    On Error GoTo Fail
    ReloadSettings
    EnsureUserConfig
    base = TextBlocksFolder()
    listing = ListTextBlocks(base)
    If Len(listing) = 0 Then extra = "<button id=""tbEmpty"" label=""" & XmlEscape(Txt("(keine Textbausteine gefunden)")) & """ enabled=""false""/>"
    extra = extra & "<button id=""tbOpenFolder"" label=""" & XmlEscape(Txt("Ordner {oe}ffnen")) & """ imageMso=""FileOpen"" onAction=""OnAction_OpenTextBlockFolder""/>" & _
        "<button id=""tbImport"" label=""" & XmlEscape(Txt("Beispielbausteine importieren" & ChrW(8230))) & """ onAction=""OnAction_ImportTextBlocks""/>"
    TextBlocksMenuContent = TextBlockMenuXml(base, listing, extra)
    Exit Function
Fail:
    TextBlocksMenuContent = TextBlockMenuXml("", "", "<button id=""tbError"" label=""" & _
        XmlEscape(Txt("Fehler: ") & Err.Description) & """ enabled=""false""/>")
End Function

' "subfolder<TAB>file" lines for the top level and one level of subfolders.
Private Function ListTextBlocks(ByVal base As String) As String
    Dim folders As New Collection, f As String, subFolder As Variant, s As String
    If Len(Dir$(base, vbDirectory)) = 0 Then Exit Function
    f = Dir$(base & "\*", vbDirectory)
    Do While Len(f) > 0
        If f <> "." And f <> ".." Then
            If (GetAttr(base & "\" & f) And vbDirectory) = vbDirectory Then
                folders.Add f
            ElseIf IsTextBlockFile(f) Then
                s = s & vbLf & vbTab & f
            End If
        End If
        f = Dir$()
    Loop
    For Each subFolder In folders
        f = Dir$(base & "\" & subFolder & "\*")
        Do While Len(f) > 0
            If IsTextBlockFile(f) Then s = s & vbLf & subFolder & vbTab & f
            f = Dir$()
        Loop
    Next subFolder
    If Len(s) > 0 Then ListTextBlocks = Mid$(s, 2)
End Function

Public Sub InsertTextBlock(ByVal path As String)
    Dim doc As Document, rng As Range, ins As Range, startPos As Long, oldLen As Long, removed As Long
    Dim names As Collection, nm As Variant, value As String, unfilled As New Collection, errDesc As String

    If Documents.Count = 0 Then
        MsgBox MsgNoDocument(), vbExclamation, AppTitle()
        Exit Sub
    End If
    If Len(Dir$(path)) = 0 Then
        MsgBox MsgTextBlockMissing(path), vbExclamation, AppTitle()
        Exit Sub
    End If
    Set doc = ActiveDocument
    ReloadSettings
    On Error GoTo Fail
    Application.UndoRecord.StartCustomRecord UndoTitleTextBlock()

    ' insert at the selection; the inserted range is derived from the length change
    ' of the story the cursor is in (main text, footer, text box, ...)
    Set rng = Selection.Range
    Set ins = rng.Duplicate
    startPos = rng.Start
    removed = rng.End - rng.Start
    oldLen = rng.StoryLength
    rng.InsertFile path
    ins.SetRange startPos, startPos + (ins.StoryLength - oldLen) + removed

    If CBool(GetSettingValue("textblocks.askForPlaceholders", True)) Then
        Set names = FindPlaceholders(ins.Text)
        For Each nm In names
            value = InputBox(MsgPlaceholderPrompt(CStr(nm)), AppTitle(), PlaceholderDefault(CStr(nm)))
            If Len(value) > 0 Then
                ReplaceInRange ins, "[" & nm & "]", value
            Else
                unfilled.Add nm
            End If
        Next nm
    Else
        Set unfilled = FindPlaceholders(ins.Text)
    End If
    If CBool(GetSettingValue("textblocks.highlightUnfilledPlaceholders", True)) Then
        For Each nm In unfilled
            HighlightInRange ins, "[" & nm & "]"
        Next nm
    End If

    Application.UndoRecord.EndCustomRecord
    ins.Select
    Selection.Collapse wdCollapseEnd
    Exit Sub
Fail:
    errDesc = Err.Description
    Application.UndoRecord.EndCustomRecord
    MsgBox MsgUnexpectedError(errDesc), vbCritical, AppTitle()
End Sub

Private Sub ReplaceInRange(ByVal ins As Range, ByVal findText As String, ByVal replaceText As String)
    Dim r As Range
    Set r = ins.Duplicate
    With r.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = findText
        .Replacement.Text = Replace(replaceText, "^", "^^")
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = False
        .Execute Replace:=wdReplaceAll
    End With
End Sub

Private Sub HighlightInRange(ByVal ins As Range, ByVal findText As String)
    Dim r As Range
    Set r = ins.Duplicate
    With r.Find
        .ClearFormatting
        .Text = findText
        .Forward = True
        .Wrap = wdFindStop
        .MatchCase = True
        .MatchWildcards = False
    End With
    Do While r.Find.Execute
        If r.End > ins.End Then Exit Do
        r.HighlightColorIndex = wdYellow
        r.Collapse wdCollapseEnd
    Loop
End Sub

Public Sub OpenTextBlockFolder()
    Dim f As String
    On Error GoTo Fail
    ReloadSettings
    EnsureUserConfig
    f = TextBlocksFolder()
    If Len(Dir$(f, vbDirectory)) = 0 Then MkDir f
    Shell "explorer.exe """ & f & """", vbNormalFocus
    Exit Sub
Fail:
    MsgBox MsgConfigFolderError(f, Err.Description), vbExclamation, AppTitle()
End Sub

' Copies the neutral example blocks (repository folder config\textblocks) into
' the text-block folder. Existing files are never overwritten.
Public Sub ImportSampleTextBlocks()
    Dim src As String, dst As String, fso As Object, subFolder As Object, n As Long
    On Error GoTo Fail
    ReloadSettings
    EnsureUserConfig
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = Txt("Ordner config\textblocks des Repositorys w{ae}hlen")
        If .Show <> -1 Then Exit Sub
        src = .SelectedItems(1)
    End With
    dst = TextBlocksFolder()
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(dst) Then fso.CreateFolder dst
    n = CopyBlocks(fso, fso.GetFolder(src), dst)
    For Each subFolder In fso.GetFolder(src).SubFolders
        If Not fso.FolderExists(dst & "\" & subFolder.Name) Then fso.CreateFolder dst & "\" & subFolder.Name
        n = n + CopyBlocks(fso, subFolder, dst & "\" & subFolder.Name)
    Next subFolder
    MsgBox MsgTextBlocksImported(n, dst), vbInformation, AppTitle()
    Exit Sub
Fail:
    MsgBox MsgUnexpectedError(Err.Description), vbCritical, AppTitle()
End Sub

Private Function CopyBlocks(ByVal fso As Object, ByVal folder As Object, ByVal dst As String) As Long
    Dim file As Object
    For Each file In folder.Files
        If IsTextBlockFile(file.Name) And Not fso.FileExists(dst & "\" & file.Name) Then
            fso.CopyFile file.Path, dst & "\" & file.Name, False
            CopyBlocks = CopyBlocks + 1
        End If
    Next file
End Function
