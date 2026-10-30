Attribute VB_Name = "HelloWorld"
Option Explicit

' Sanity-check macro: confirms macros are enabled and the module imported correctly.
Public Sub HelloWorld()
    MsgBox "Macros are working in " & ActiveDocument.Name, vbInformation
End Sub
