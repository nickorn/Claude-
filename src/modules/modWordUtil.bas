Attribute VB_Name = "modWordUtil"
Option Explicit

' Shared Word helpers: tool comments and the revision view used while reading text.

' Adds a comment authored by the tool (settings: check.commentAuthor), so that
' "Pruefmarkierungen entfernen" can later remove exactly these comments.
Public Sub AddToolComment(ByVal doc As Document, ByVal startPos As Long, ByVal endPos As Long, ByVal commentText As String)
    Dim author As String, oldName As String, oldInitials As String, c As Comment
    Dim errNum As Long, errDesc As String
    author = CStr(GetSettingValue("check.commentAuthor", "PatentCheck"))
    oldName = Application.UserName
    oldInitials = Application.UserInitials
    Application.UserName = author
    Application.UserInitials = "PC"

    On Error Resume Next
    Set c = doc.Comments.Add(doc.Range(startPos, endPos), commentText)
    errNum = Err.Number
    errDesc = Err.Description
    If errNum = 0 Then
        c.Author = author
        c.Initial = "PC"
    End If
    Err.Clear
    On Error GoTo 0

    Application.UserName = oldName
    Application.UserInitials = oldInitials
    If errNum <> 0 Then Err.Raise errNum, "AddToolComment", errDesc
End Sub

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
