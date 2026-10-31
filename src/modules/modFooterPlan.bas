Attribute VB_Name = "modFooterPlan"
Option Explicit

' Where the status text goes in a footer paragraph (pure, no Word objects).
' Layout convention: status bottom left, application number bottom right,
' separated by a TAB (or in a two-cell table, handled by modVersions).
' The right part - the application number typed by hand - is NEVER changed.

Public Const FA_NONE As String = "none"
Public Const FA_REPLACE As String = "replace"          ' replace [start, start+len) with text
Public Const FA_REPLACE_EMPTY As String = "replaceEmpty" ' empty footer: text + TAB, right tab stop needed
Public Const FA_INSERT_BEFORE As String = "insertBefore" ' new paragraph with text before the footer paragraph

' Returns Array(action, start, length, text).
'   paraText     footer paragraph without paragraph mark
'   status       new status text ("" = remove the status)
'   knownStatus  "|Reinschrift|Clean version|...|" all status texts of all languages
Public Function PlanFooterLeft(ByVal paraText As String, ByVal status As String, ByVal knownStatus As String) As Variant
    Dim q As Long
    q = InStr(1, paraText, vbTab)
    If q > 0 Then
        ' "status<TAB>application number": only the part before the first TAB changes
        If Left$(paraText, q - 1) = status Then
            PlanFooterLeft = Array(FA_NONE, 0, 0, "")
        Else
            PlanFooterLeft = Array(FA_REPLACE, 1, q - 1, status)
        End If
    ElseIf Len(Trim$(paraText)) = 0 Then
        If Len(status) = 0 Then
            PlanFooterLeft = Array(FA_NONE, 0, 0, "")
        Else
            PlanFooterLeft = Array(FA_REPLACE_EMPTY, 1, Len(paraText), status & vbTab)
        End If
    ElseIf InStr(1, knownStatus, "|" & Trim$(paraText) & "|") > 0 Then
        ' the paragraph holds only an earlier status
        If Trim$(paraText) = status Then
            PlanFooterLeft = Array(FA_NONE, 0, 0, "")
        Else
            PlanFooterLeft = Array(FA_REPLACE, 1, Len(paraText), status)
        End If
    ElseIf Len(status) = 0 Then
        PlanFooterLeft = Array(FA_NONE, 0, 0, "")
    Else
        ' unknown content without TAB (e.g. only the application number): keep it untouched
        PlanFooterLeft = Array(FA_INSERT_BEFORE, 0, 0, status)
    End If
End Function
