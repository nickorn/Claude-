Attribute VB_Name = "modTestFormat"
Option Explicit

' Unit tests for the MVP-3 text fixes (modFormatText).
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunFormatTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunFormatTests()
    MsgBox RunFormatTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunFormatTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    TestSpaces
    TestPunctuationEnglish
    TestPunctuationGerman
    TestPunctuationLeavesInnerStop
    TestSigns
    TestFixThenCheck
    RunFormatTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

Private Function L(ParamArray lines() As Variant) As Variant
    Dim arr() As String, i As Long
    ReDim arr(0 To UBound(lines) - LBound(lines))
    For i = LBound(lines) To UBound(lines)
        arr(i - LBound(lines)) = Replace(CStr(lines(i)), "|", vbTab)
    Next i
    L = arr
End Function

Private Sub Check(ByVal name As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If cond Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & name & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
    End If
End Sub

' kind: 1 spaces, 2 punctuation, 3 signs
Private Function ApplyFix(ByVal lines As Variant, ByVal kind As Long) As Variant
    Dim ops As New Collection, flagLog As String
    SetDefaultCheckOptions
    BuildClaimModel lines, Empty, "auto"
    Select Case kind
        Case 1: PlanSpaceFixes ops
        Case 2: PlanPunctuationFixes ops
        Case 3: PlanSignFormatFixes ops
    End Select
    ApplyFix = ApplyOpsToTexts(lines, ops, flagLog)
End Function

Private Sub ExpectLines(ByVal name As String, ByVal got As Variant, ByVal expected As Variant)
    Dim i As Long
    For i = LBound(expected) To UBound(expected)
        Check name & " line " & i, got(i) = expected(i), "got [" & Replace(got(i), vbTab, "|") & "] expected [" & _
            Replace(expected(i), vbTab, "|") & "]"
    Next i
End Sub

Private Sub TestSpaces()
    ExpectLines "spaces", ApplyFix(L("Two  spaces , heading", "1.|A device  with   a part , and a lens .  ", "", "2.|Wait ... more."), 1), _
        L("Two spaces , heading", "1.|A device with a part, and a lens.", "", "2.|Wait ... more.")
End Sub

Private Sub TestPunctuationEnglish()
    ExpectLines "punct en", ApplyFix(L("1.|A method comprising", "receiving a signal", "filtering the signal,", "outputting the signal;"), 2), _
        L("1.|A method comprising:", "receiving a signal;", "filtering the signal; and", "outputting the signal.")
    ExpectLines "punct double stop", ApplyFix(L("1.|A device..", "", "2.|The device of claim 1"), 2), _
        L("1.|A device.", "", "2.|The device of claim 1.")
    ExpectLines "punct keeps ok", ApplyFix(L("1.|A method comprising:", "x;", "y; and", "z."), 2), _
        L("1.|A method comprising:", "x;", "y; and", "z.")
End Sub

Private Sub TestPunctuationGerman()
    ExpectLines "punct de", ApplyFix(L("1.|Verfahren mit:", "Empfangen eines Signals;", "Filtern des Signals;", "Ausgeben des Signals"), 2), _
        L("1.|Verfahren mit:", "Empfangen eines Signals;", "Filtern des Signals; und", "Ausgeben des Signals.")
End Sub

Private Sub TestPunctuationLeavesInnerStop()
    ' text after an inner full stop: only the list up to the stop is fixed
    ExpectLines "punct inner stop", ApplyFix(L("1.|A method comprising:", "x;", "y.", "wherein z;"), 2), _
        L("1.|A method comprising:", "x; and", "y.", "wherein z;")
End Sub

Private Sub TestSigns()
    ExpectLines "signs", ApplyFix(L("1.|A controller(150;450) and a sensor ( 2 , 3 ) and (a) step (150, 450)."), 3), _
        L("1.|A controller (150, 450) and a sensor (2, 3) and (a) step (150, 450).")
End Sub

Private Sub TestFixThenCheck()
    Dim lines As Variant, f As Collection
    lines = L("1.|Method for authorizing a third party, the method comprising:", "", _
        "performing, at an access controller:", "", _
        "receiving (1210) a secret value; and", "", _
        "receiving (1220) a combined secret value;", "", _
        "processing (1230) the combined secret value  using the secret value", "", _
        "authorizing (1330) the third party", "", _
        "2.|The method of claim 1, wherein the comparing is executed at the access controller (150; 450) .")
    lines = ApplyFix(lines, 1)
    lines = ApplyFix(lines, 2)
    lines = ApplyFix(lines, 3)
    SetDefaultCheckOptions
    BuildClaimModel lines, Empty, "auto"
    Set f = New Collection
    CheckSpacing f
    Check "fixed spacing", f.Count = 0, CStr(f.Count)
    Set f = New Collection
    CheckPunctuation f
    ' only the misplaced "; and" (1210) remains - deliberately not changed automatically
    Check "fixed punctuation", f.Count = 1, CStr(f.Count)
    Set f = New Collection
    CheckReferenceSigns f
    ' the semicolon is gone; the missing sign for "an access controller" in claim 1 remains
    Check "fixed signs", f.Count = 1, CStr(f.Count)
End Sub
