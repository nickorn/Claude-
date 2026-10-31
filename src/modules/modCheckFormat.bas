Attribute VB_Name = "modCheckFormat"
Option Explicit

' Formatting check (needs Word paragraph properties): font, size, line spacing
' and space before/after against settings.json "format".

Public Sub CheckFormatting(ByVal doc As Document, ByVal texts As Variant, ByVal findings As Collection)
    Dim para As Paragraph, i As Long, details As String, flagged As Long, firstExtra As Long
    Dim fontName As String, fontSize As Single, lineSpacing As Single
    Dim spaceBefore As Single, spaceAfter As Single, maxFindings As Long, wholeDoc As Boolean
    Dim f As Font, startPara As Long

    fontName = CStr(GetSettingValue("format.fontName", "Arial"))
    fontSize = CSng(GetSettingValue("format.fontSize", 11))
    lineSpacing = CSng(GetSettingValue("format.lineSpacing", 1.5))
    spaceBefore = CSng(GetSettingValue("format.spaceBefore", 0))
    spaceAfter = CSng(GetSettingValue("format.spaceAfter", 0))
    maxFindings = CLng(GetSettingValue("check.maxFormattingFindings", 20))
    wholeDoc = CBool(GetSettingValue("format.applyToWholeDocument", True))
    If Not wholeDoc And gClaimCount > 0 Then startPara = gClaimPara(1)
    firstExtra = -1

    For Each para In doc.Paragraphs
        If i >= startPara And Len(Trim$(texts(i))) > 0 Then
            details = ""
            Set f = para.Range.Font
            If f.Name <> fontName Then details = AppendDetail(details, MsgFormatFont(f.Name, fontName))
            If f.Size <> fontSize Then
                If f.Size >= 9999 Then
                    details = AppendDetail(details, MsgFormatSize("gemischt", CStr(fontSize)))
                Else
                    details = AppendDetail(details, MsgFormatSize(CStr(f.Size), CStr(fontSize)))
                End If
            End If
            If Not LineSpacingMatches(para, lineSpacing) Then
                details = AppendDetail(details, MsgFormatLineSpacing(CStr(lineSpacing)))
            End If
            If para.SpaceBefore <> spaceBefore Or para.SpaceAfter <> spaceAfter Then
                details = AppendDetail(details, MsgFormatParaSpacing())
            End If
            If Len(details) > 0 Then
                flagged = flagged + 1
                If flagged <= maxFindings Then
                    AddFinding findings, "formatting", i, 1, Len(texts(i)), MsgFormatParagraph(details)
                ElseIf firstExtra < 0 Then
                    firstExtra = i
                End If
            End If
        End If
        i = i + 1
    Next para
    If flagged > maxFindings Then
        AddFinding findings, "formatting", firstExtra, 1, 0, MsgFormatMore(flagged - maxFindings)
    End If
End Sub

Private Function AppendDetail(ByVal details As String, ByVal item As String) As String
    If Len(details) > 0 Then details = details & ", "
    AppendDetail = details & item
End Function

Public Function LineSpacingMatches(ByVal para As Paragraph, ByVal factor As Single) As Boolean
    Select Case para.LineSpacingRule
        Case wdLineSpaceSingle
            LineSpacingMatches = (Abs(factor - 1) < 0.01)
        Case wdLineSpace1pt5
            LineSpacingMatches = (Abs(factor - 1.5) < 0.01)
        Case wdLineSpaceDouble
            LineSpacingMatches = (Abs(factor - 2) < 0.01)
        Case wdLineSpaceMultiple
            LineSpacingMatches = (Abs(para.LineSpacing - LinesToPoints(factor)) < 0.5)
    End Select
End Function
