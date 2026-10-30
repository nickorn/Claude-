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

Public Sub OnAction_OpenSettings(control As IRibbonControl)
    OpenSettingsFile
End Sub

Public Sub OnAction_Info(control As IRibbonControl)
    MsgBox MsgInfo(ConfigFolder()), vbInformation, AppTitle()
End Sub
