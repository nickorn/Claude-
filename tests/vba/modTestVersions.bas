Attribute VB_Name = "modTestVersions"
Option Explicit

' Unit tests for the footer status planning (modFooterPlan).
' The right part of the footer (application number) must never change.
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunVersionTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunVersionTests()
    MsgBox RunVersionTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunVersionTests() As String
    Dim known As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    known = "|Markierte Fassung|Reinschrift|Vom Erfinder freigegeben|Marked-up version|Clean version|Approved by inventor|"

    ExpectPlan "tab layout replace", "Markierte Fassung" & vbTab & "20 187 485.6", "Reinschrift", known, _
        "Reinschrift" & vbTab & "20 187 485.6"
    ExpectPlan "tab layout empty left", vbTab & "20 187 485.6", "Clean version", known, "Clean version" & vbTab & "20 187 485.6"
    ExpectPlan "tab layout unchanged", "Reinschrift" & vbTab & "EP 1", "Reinschrift", known, "Reinschrift" & vbTab & "EP 1"
    ExpectPlan "tab layout remove", "Reinschrift" & vbTab & "EP 1", "", known, vbTab & "EP 1"
    ExpectPlan "empty footer", "", "Reinschrift", known, "Reinschrift" & vbTab
    ExpectPlan "old status only", "Marked-up version", "Clean version", known, "Clean version"
    ExpectPlan "remove old status only", "Reinschrift", "", known, ""
    ' only the application number, no TAB: it stays untouched, status goes into a new paragraph
    ExpectAction "number only", "20 187 485.6", "Reinschrift", known, FA_INSERT_BEFORE
    ExpectAction "number only remove", "20 187 485.6", "", known, FA_NONE
    ExpectAction "empty remove", "", "", known, FA_NONE

    RunVersionTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

Private Sub Check(ByVal testName As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If cond Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & testName & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
    End If
End Sub

' Applies the plan to the text and compares.
Private Sub ExpectPlan(ByVal testName As String, ByVal t As String, ByVal status As String, ByVal known As String, ByVal expected As String)
    Dim plan As Variant, got As String
    plan = PlanFooterLeft(t, status, known)
    Select Case plan(0)
        Case FA_NONE
            got = t
        Case FA_REPLACE, FA_REPLACE_EMPTY
            got = Left$(t, plan(1) - 1) & plan(3) & Mid$(t, plan(1) + plan(2))
        Case Else
            got = "<" & plan(0) & ">"
    End Select
    Check testName, got = expected, "got [" & Replace(got, vbTab, "|") & "] expected [" & Replace(expected, vbTab, "|") & "]"
    ' whatever happens, the text after the first TAB is preserved
    If InStr(t, vbTab) > 0 Then Check testName & " right part kept", Mid$(got, InStr(got, vbTab)) = Mid$(t, InStr(t, vbTab))
End Sub

Private Sub ExpectAction(ByVal testName As String, ByVal t As String, ByVal status As String, ByVal known As String, ByVal action As String)
    Dim plan As Variant
    plan = PlanFooterLeft(t, status, known)
    Check testName, plan(0) = action, "got " & plan(0) & " expected " & action
End Sub
