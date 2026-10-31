Attribute VB_Name = "modTestDraft"
Option Explicit

' Unit tests for the drafting helpers (modDraftPlan): claim tree, dependent
' claim skeleton, merge planning.
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunDraftTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunDraftTests()
    MsgBox RunDraftTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunDraftTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    TestTree
    TestSkeleton
    TestMergeWherein
    TestMergeFeatures
    TestMergeRefused
    RunDraftTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

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

Private Sub Model(ByVal lines As Variant)
    SetDefaultCheckOptions
    BuildClaimModel lines, Empty, "auto"
End Sub

Private Sub TestTree()
    Dim t As Collection, s As String, ln As Variant
    Model L("1.|A method comprising a step.", "", "2.|The method of claim 1, wherein x.", "", _
        "3.|The method of claim 2, wherein y.", "", "4.|The method of claim 1 or 3, wherein z.", "", _
        "5.|A device for carrying out the method of claim 1.", "", "6.|A system comprising a server.")
    Set t = ClaimTreeLines()
    For Each ln In t
        s = s & ln(TL_DEPTH) & ":" & ln(TL_LABEL) & " "
    Next ln
    Check "tree order and depth", s = "0:1 1:2 2:3 1:4 1:5 0:6 ", s
    Check "tree note multi", InStr(t(4)(TL_NOTE), "3") > 0, CStr(t(4)(TL_NOTE))
    Check "tree note category", t(1)(TL_NOTE) = "Verfahren", CStr(t(1)(TL_NOTE))
End Sub

Private Sub TestSkeleton()
    Model L("1.|A method for filtering a signal.", "", "2.|The method of claim 1, wherein x.", "", "2a.|The method of claim 2, wherein y.")
    Check "skeleton en", DependentSkeleton(2) = "The method according to claim 2, wherein ", DependentSkeleton(2)
    Check "next label", NextSupplementaryLabel("2") = "2b", NextSupplementaryLabel("2")
    Check "next label free", NextSupplementaryLabel("1") = "1a", NextSupplementaryLabel("1")
    Model L("1.|Steuerungsvorrichtung f" & ChrW(252) & "r ein Fahrzeug, mit einem Sensor.", "", "2.|Steuerungsvorrichtung nach Anspruch 1, wobei x.")
    Check "skeleton de", DependentSkeleton(2) = "Steuerungsvorrichtung nach Anspruch 2, wobei ", DependentSkeleton(2)
    Model L("1.|A computer-readable storage medium storing a program.")
    Check "skeleton medium", DependentSkeleton(1) = "The medium according to claim 1, wherein ", DependentSkeleton(1)
    Model L("1.|Computer program product comprising instructions.")
    Check "skeleton program", DependentSkeleton(1) = "The program according to claim 1, wherein ", DependentSkeleton(1)
End Sub

' Applies the merge plan to the parent's texts (as modDraft does in Word).
Private Function MergedParent(ByVal lines As Variant, ByVal d As Long) As String
    Dim plan As Variant, last As String, p As Long, s As String, pen As String
    Model lines
    plan = PlanMerge(d)
    If Not plan(0) Then
        MergedParent = "REFUSED: " & plan(1)
        Exit Function
    End If
    For p = gClaimPara(plan(2)) To gClaimLastPara(plan(2))
        If p = gClaimLastPara(plan(2)) Then
            last = RTrimText(gTexts(p))
            s = s & Left$(last, Len(last) - 1) & plan(3)
        ElseIf p = plan(4) Then
            pen = RTrimText(gTexts(p))
            s = s & Left$(pen, Len(pen) - plan(5)) & "/"
        ElseIf Not IsBlankText(gTexts(p)) Then
            s = s & gTexts(p) & "/"
        End If
    Next p
    MergedParent = Replace(Replace(s, vbCr, "/"), vbTab, "|")
End Function

Private Sub TestMergeWherein()
    Dim got As String
    got = MergedParent(L("1.|A device comprising a sensor and a lens.", "", "2.|The device of claim 1, wherein a sensor is optical and a filter is red."), 2)
    Check "merge wherein + articles", got = "1.|A device comprising a sensor and a lens, wherein the sensor is optical and a filter is red.", got
End Sub

Private Sub TestMergeFeatures()
    Dim got As String
    got = MergedParent(L("1.|A method comprising:", "receiving a signal; and", "filtering the signal.", "", _
        "2.|The method of claim 1, further comprising storing a signal."), 2)
    Check "merge feature + and moved", got = "1.|A method comprising:/receiving a signal;/filtering the signal; and/storing the signal.", got
    got = MergedParent(L("1.|Verfahren mit:", "Empfangen eines Signals; und", "Filtern des Signals.", "", _
        "2.|Verfahren nach Anspruch 1, ferner umfassend:", "Speichern des Signals."), 2)
    Check "merge de list", got = "1.|Verfahren mit:/Empfangen eines Signals;/Filtern des Signals; und/Speichern des Signals.", got
End Sub

Private Sub TestMergeRefused()
    Check "merge independent refused", Left$(MergedParent(L("1.|A device."), 1), 8) = "REFUSED:"
    Check "merge multi refused", Left$(MergedParent(L("1.|A.", "", "2.|B.", "", "3.|The B of claim 1 or 2, wherein x."), 3), 8) = "REFUSED:"
    Check "merge parent no stop", Left$(MergedParent(L("1.|A device", "", "2.|The device of claim 1, wherein x."), 2), 8) = "REFUSED:"
End Sub
