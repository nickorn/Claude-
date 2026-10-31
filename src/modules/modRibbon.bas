Attribute VB_Name = "modRibbon"
Option Explicit

' Ribbon callbacks for src/customUI/customUI14.xml.

Private mRibbon As IRibbonUI

Public Sub OnRibbonLoad(ribbon As IRibbonUI)
    Set mRibbon = ribbon
    ' first run: create %APPDATA%\PatentTools with the default files
    On Error Resume Next
    EnsureUserConfig
End Sub

Public Sub OnAction_Renumber(control As IRibbonControl)
    RenumberClaims
End Sub

Public Sub OnAction_FullCheck(control As IRibbonControl)
    RunFullCheck
End Sub

Public Sub OnAction_CheckSingle(control As IRibbonControl)
    RunSingleCheck control.Tag
End Sub

Public Sub OnAction_RemoveMarks(control As IRibbonControl)
    RemoveCheckMarks
End Sub

Public Sub OnAction_Report(control As IRibbonControl)
    RunCheckReport
End Sub

Public Sub OnAction_ApplyFormat(control As IRibbonControl)
    ApplyFormatting
End Sub

Public Sub OnAction_FixSpaces(control As IRibbonControl)
    FixSpaces
End Sub

Public Sub OnAction_FixPunctuation(control As IRibbonControl)
    FixPunctuation
End Sub

Public Sub OnAction_FixSigns(control As IRibbonControl)
    FixSignFormat
End Sub

Public Sub OnAction_CleanVersion(control As IRibbonControl)
    CreateCleanVersion
End Sub

Public Sub OnAction_MarkedUpVersion(control As IRibbonControl)
    CreateMarkedUpVersion
End Sub

Public Sub OnAction_FooterStatus(control As IRibbonControl)
    SetFooterStatus control.Tag
End Sub

Public Sub GetContent_TextBlocks(control As IRibbonControl, ByRef returnedVal)
    returnedVal = TextBlocksMenuContent()
End Sub

Public Sub OnAction_InsertTextBlock(control As IRibbonControl)
    InsertTextBlock control.Tag
End Sub

Public Sub OnAction_OpenTextBlockFolder(control As IRibbonControl)
    OpenTextBlockFolder
End Sub

Public Sub OnAction_ImportTextBlocks(control As IRibbonControl)
    ImportSampleTextBlocks
End Sub

Public Sub OnAction_OpenSettings(control As IRibbonControl)
    OpenSettingsFile
End Sub

Public Sub OnAction_Info(control As IRibbonControl)
    MsgBox MsgInfo(ConfigFolder()), vbInformation, AppTitle()
End Sub
