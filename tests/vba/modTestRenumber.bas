Attribute VB_Name = "modTestRenumber"
Option Explicit

' Unit tests for modClaimParse (pure text logic of the renumbering).
' Not part of PatentTools.dotm. To run in Word: import this module next to
' the src modules, then Alt+F8 -> PT_RunRenumberTests.
' Also runs headless in LibreOffice Basic via tools/run_vba_tests.py.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunRenumberTests()
    MsgBox RunRenumberTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunRenumberTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0

    TestSupplementaryEnglish
    TestGermanReferences
    TestOpaqueLabels
    TestPositionOrder
    TestDuplicates
    TestUnknownReference
    TestKeywordBoundaries
    TestRangesAndDashes
    TestTextBeforeFirstClaimIgnored
    TestAlreadyConsecutive
    TestNoClaims
    TestCaseSensitiveLabels
    TestClaimLabelGrammar

    RunRenumberTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

' ---------------------------------------------------------------- helpers

Private Function L(ParamArray lines() As Variant) As Variant
    Dim arr() As String, i As Long
    ReDim arr(0 To UBound(lines) - LBound(lines))
    For i = LBound(lines) To UBound(lines)
        arr(i - LBound(lines)) = Replace(CStr(lines(i)), "|", vbTab)
    Next i
    L = arr
End Function

Private Sub Check(ByVal testName As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If cond Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & testName & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
    End If
End Sub

Private Function Visible(ByVal s As String) As String
    Visible = Replace(s, vbTab, "|")
End Function

' Runs the planner and compares every paragraph with expected (| = TAB).
Private Sub ExpectRenumber(ByVal testName As String, ByVal inputLines As Variant, _
        ByVal expectedLines As Variant, ByVal expectedFlags As Long)
    Dim ops As New Collection, errMsg As String, n As Long, outLines As Variant
    Dim flagLog As String, i As Long, flags As Long
    If Not PlanRenumber(inputLines, ops, errMsg, n) Then
        Check testName & " (plan)", False, errMsg
        Exit Sub
    End If
    outLines = ApplyOpsToTexts(inputLines, ops, flagLog)
    For i = LBound(expectedLines) To UBound(expectedLines)
        Check testName & " line " & i, outLines(i) = expectedLines(i), _
            "got [" & Visible(CStr(outLines(i))) & "] expected [" & Visible(CStr(expectedLines(i))) & "]"
    Next i
    If Len(flagLog) > 0 Then flags = UBound(Split(flagLog, vbLf))
    Check testName & " flags", flags = expectedFlags, "got " & flags & " expected " & expectedFlags & vbLf & flagLog
End Sub

' ---------------------------------------------------------------- tests

Private Sub TestSupplementaryEnglish()
    ExpectRenumber "EN supplementary", _
        L("Claims", "", _
          "1.|A method comprising a step.", "", _
          "1a.|The method of claim 1, wherein x.", "", _
          "2.|The method of claim 1 or 1a, wherein y.", "", _
          "3.|A device configured to carry out the method of any one of claims 1 to 2."), _
        L("Claims", "", _
          "1.|A method comprising a step.", "", _
          "2.|The method of claim 1, wherein x.", "", _
          "3.|The method of claim 1 or 2, wherein y.", "", _
          "4.|A device configured to carry out the method of any one of claims 1 to 3."), 1
End Sub

Private Sub TestGermanReferences()
    Dim ue As String
    ue = ChrW(252)
    ExpectRenumber "DE references", _
        L("Patentanspr" & ue & "che", _
          "1.|Verfahren mit einem Schritt.", "", _
          "1.1.|Verfahren nach Anspruch 1, wobei x.", "", _
          "2.|Verfahren nach einem der Anspr" & ue & "che 1, 1.1 oder 2a, wobei y.", "", _
          "2a.|Verfahren gem" & ChrW(228) & ChrW(223) & " Patentanspruch 1.1.", "", _
          "3.|Vorrichtung zur Durchf" & ue & "hrung des Verfahrens nach den Anspr" & ue & "chen 1 bis 2a sowie des Anspruchs 2."), _
        L("Patentanspr" & ue & "che", _
          "1.|Verfahren mit einem Schritt.", "", _
          "2.|Verfahren nach Anspruch 1, wobei x.", "", _
          "3.|Verfahren nach einem der Anspr" & ue & "che 1, 2 oder 4, wobei y.", "", _
          "4.|Verfahren gem" & ChrW(228) & ChrW(223) & " Patentanspruch 2.", "", _
          "5.|Vorrichtung zur Durchf" & ue & "hrung des Verfahrens nach den Anspr" & ue & "chen 1 bis 4 sowie des Anspruchs 3."), 0
End Sub

Private Sub TestOpaqueLabels()
    ' 1a1 and 1a.1 are different claims
    ExpectRenumber "opaque labels", _
        L("1.|A.", "", "1a1.|Claim 1.", "", "1a.1.|Claims 1a1 and 1a.1.", "", "12b.3.|Claim 1a.1."), _
        L("1.|A.", "", "2.|Claim 1.", "", "3.|Claims 2 and 3.", "", "4.|Claim 3."), 0
End Sub

Private Sub TestPositionOrder()
    ' new order = document order, not numeric order
    ExpectRenumber "position order", _
        L("1.|A.", "", "3.|The method of claim 1.", "", "2.|The method of claim 3."), _
        L("1.|A.", "", "2.|The method of claim 1.", "", "3.|The method of claim 2."), 0
End Sub

Private Sub TestDuplicates()
    Dim ops As New Collection, errMsg As String, n As Long
    Check "duplicates abort", Not PlanRenumber(L("1.|A.", "", "2.|B.", "", "2.|C.", "", "1.|D."), ops, errMsg, n)
    Check "duplicates listed", InStr(errMsg, "2, 1") > 0, errMsg
    Check "duplicates no ops", ops.Count = 0
End Sub

Private Sub TestUnknownReference()
    ExpectRenumber "unknown reference", _
        L("1.|A.", "", "1a.|The method of claim 7 or 1."), _
        L("1.|A.", "", "2.|The method of claim 7 or 1."), 1
End Sub

Private Sub TestKeywordBoundaries()
    ' "claimed", "disclaim", "Anspruchsbaum" are no keywords; trailing full stop is not part of the label
    ExpectRenumber "keyword boundaries", _
        L("1.|A.", "", "1a.|As claimed 1a times, disclaim 1a, Anspruchsbaum 1a, see claim 1a.", "", "2.|Claim 1a; claim 1a, and claim 2."), _
        L("1.|A.", "", "2.|As claimed 1a times, disclaim 1a, Anspruchsbaum 1a, see claim 2.", "", "3.|Claim 2; claim 2, and claim 3."), 0
End Sub

Private Sub TestRangesAndDashes()
    ' contiguous ranges stay silent; dash ranges and and/or lists are rewritten
    ExpectRenumber "ranges", _
        L("1a.|A.", "", "1b.|B.", "", "1c.|C.", "", "1d.|Claims 1a-1c and/or 1b " & ChrW(8211) & " 1c, claims 1a through 1b."), _
        L("1.|A.", "", "2.|B.", "", "3.|C.", "", "4.|Claims 1-3 and/or 2 " & ChrW(8211) & " 3, claims 1 through 2."), 0
    ' reversed range after moving a claim -> flagged
    ExpectRenumber "reversed range", _
        L("1.|A.", "", "3.|B.", "", "2.|C.", "", "4.|Claims 2 to 3."), _
        L("1.|A.", "", "2.|B.", "", "3.|C.", "", "4.|Claims 3 to 2."), 1
End Sub

Private Sub TestTextBeforeFirstClaimIgnored()
    ExpectRenumber "preamble ignored", _
        L("Claims 1a to 2 as amended", "", "1a.|A.", "", "2.|Claim 1a."), _
        L("Claims 1a to 2 as amended", "", "1.|A.", "", "2.|Claim 1."), 0
End Sub

Private Sub TestAlreadyConsecutive()
    Dim ops As New Collection, errMsg As String, n As Long
    Check "consecutive plan", PlanRenumber(L("Claims", "1.|A.", "", "2.|Claim 1.", "", "3.|Claims 1 to 2."), ops, errMsg, n)
    Check "consecutive count", n = 3, CStr(n)
    Check "consecutive no ops", ops.Count = 0, CStr(ops.Count)
End Sub

Private Sub TestNoClaims()
    Dim ops As New Collection, errMsg As String, n As Long
    Check "no claims", Not PlanRenumber(L("Claims", "1 A.", "1.A.", "a.|B."), ops, errMsg, n)
    Check "no claims message", Len(errMsg) > 0
End Sub

Private Sub TestCaseSensitiveLabels()
    ExpectRenumber "case-sensitive labels", _
        L("1.|A.", "", "1a.|Claim 1.", "", "1A.|Claims 1a, 1A."), _
        L("1.|A.", "", "2.|Claim 1.", "", "3.|Claims 2, 3."), 0
End Sub

Private Sub TestClaimLabelGrammar()
    Dim lbl As String, b As Long
    Check "label 12b.3", ParseClaimLabel("12b.3." & vbTab & "X", lbl, b) And lbl = "12b.3" And b = 8, lbl & "/" & b
    Check "label needs tab", Not ParseClaimLabel("1. X", lbl, b)
    Check "label needs dot", Not ParseClaimLabel("1" & vbTab & "X", lbl, b)
    Check "label starts with digit", Not ParseClaimLabel("a1." & vbTab & "X", lbl, b)
    Check "label no space", Not ParseClaimLabel("1 a." & vbTab & "X", lbl, b)
End Sub
