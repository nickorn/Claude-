Attribute VB_Name = "modTestConvert"
Option Explicit

' Unit tests for roadmap 7 (modConvertPlan): verb forms, method <-> apparatus,
' claims -> description, reference signs strip / insert.
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunConvertTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunConvertTests()
    MsgBox RunConvertTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunConvertTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    TestVerbForms
    TestMethodToApparatus
    TestApparatusToMethod
    TestGerman
    TestDescription
    TestSigns
    RunConvertTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
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

Private Function Joined(ByVal res As Variant) As String
    If Not res(0) Then
        Joined = "REFUSED: " & res(1)
    Else
        Joined = Join(res(2), "/")
    End If
End Function

Private Sub ExpectEq(ByVal testName As String, ByVal got As String, ByVal expected As String)
    Check testName, got = expected, "got [" & got & "] expected [" & expected & "]"
End Sub

Private Sub TestVerbForms()
    ExpectEq "g2v obtaining", GerundToVerb("obtaining"), "obtain"
    ExpectEq "g2v receiving", GerundToVerb("receiving"), "receive"
    ExpectEq "g2v generating", GerundToVerb("generating"), "generate"
    ExpectEq "g2v stopping", GerundToVerb("stopping"), "stop"
    ExpectEq "g2v Comparing", GerundToVerb("Comparing"), "Compare"
    ExpectEq "g2v analysing", GerundToVerb("analysing"), "analyse"
    ExpectEq "g2v not gerund", GerundToVerb("thing"), "thing"
    ExpectEq "v2g receive", VerbToGerund("receive"), "receiving"
    ExpectEq "v2g transmit", VerbToGerund("transmit"), "transmitting"
    ExpectEq "v2g obtain", VerbToGerund("obtain"), "obtaining"
    ExpectEq "v2g tie", VerbToGerund("tie"), "tying"
    ExpectEq "replace word", ReplaceWord("The method and methods of a Method", "method", "apparatus"), "The apparatus and methods of a Apparatus"
End Sub

Private Sub TestMethodToApparatus()
    Model L("1.|Method for authorizing a third party, the method comprising:", "obtaining (1110) a biometric feature;", _
        "", "the method further comprising:", "comparing the feature; and", "authorizing the third party.")
    ExpectEq "m2a", Joined(ConvertClaimCategory(1)), _
        "Apparatus for authorizing a third party, the apparatus being configured to:/obtain (1110) a biometric feature;//" & _
        "the apparatus being further configured to:/compare the feature; and/authorize the third party."
    Model L("1.|A method of filtering, comprising:", "receiving a signal.")
    ExpectEq "m2a article", Joined(ConvertClaimCategory(1)), "An apparatus of filtering, configured to:/receive a signal."
    Model L("1.|A method.", "", "2.|The method of claim 1, wherein x.")
    Check "m2a dependent refused", Left$(Joined(ConvertClaimCategory(2)), 8) = "REFUSED:"
End Sub

Private Sub TestApparatusToMethod()
    Model L("1.|A device for measuring a signal, comprising:", "a sensor configured to detect the signal;", "a processor (20) connected to the sensor; and", _
        "a display configured to show the result.")
    ExpectEq "a2m", Joined(ConvertClaimCategory(1)), _
        "A method for measuring a signal, comprising:/detecting the signal;/a processor (20) connected to the sensor; and/showing the result."
    Model L("1.|An apparatus configured to:", "receive a signal; and", "store the signal.")
    ExpectEq "a2m configured list", Joined(ConvertClaimCategory(1)), "A method, comprising:/receiving a signal; and/storing the signal."
End Sub

Private Sub TestGerman()
    Model L("1.|Verfahren zum Messen eines Signals, mit:", "Empfangen eines Signals; und", "Filtern des Signals.")
    ExpectEq "de m2a", Joined(ConvertClaimCategory(1)), _
        "Vorrichtung zum Messen eines Signals, die eingerichtet ist zum:/Empfangen eines Signals; und/Filtern des Signals."
    Model L("1.|Vorrichtung zum Messen eines Signals, mit:", "einer Einrichtung zum Empfangen eines Signals; und", "einer Einheit zum Filtern des Signals.")
    ExpectEq "de a2m", Joined(ConvertClaimCategory(1)), _
        "Verfahren zum Messen eines Signals, mit:/Empfangen eines Signals; und/Filtern des Signals."
End Sub

Private Sub TestDescription()
    Model L("1.|Method for authorizing a third party, the method comprising:", "obtaining a feature; and", "authorizing the third party.", "", _
        "2.|The method of claim 1, wherein the feature is biometric.", "", "3.|The method of claim 1, further comprising storing the feature.")
    ExpectEq "desc en independent", ClaimToDescription(1, "According to an embodiment, "), _
        "According to an embodiment, a method for authorizing a third party is provided, the method comprising: obtaining a feature; and authorizing the third party."
    ExpectEq "desc en wherein", ClaimToDescription(2, "According to an embodiment, "), "According to an embodiment, the feature is biometric."
    ExpectEq "desc en further", ClaimToDescription(3, "In examples, "), "In examples, the method further comprises storing the feature."
    Model L("1.|Vorrichtung zum Messen eines Signals, mit einem Sensor.", "", "2.|Vorrichtung nach Anspruch 1, wobei der Sensor optisch ist.")
    ExpectEq "desc de independent", ClaimToDescription(1, "Gem" & ChrW(228) & ChrW(223) & " einer Ausf" & ChrW(252) & "hrungsform "), _
        "Gem" & ChrW(228) & ChrW(223) & " einer Ausf" & ChrW(252) & "hrungsform wird eine Vorrichtung zum Messen eines Signals bereitgestellt, mit einem Sensor."
    ExpectEq "desc de wobei", ClaimToDescription(2, "In Beispielen "), "In Beispielen ist vorgesehen, dass der Sensor optisch ist."
End Sub

Private Sub TestSigns()
    Dim ops As New Collection, flagLog As String, out As Variant, map As Collection
    ' (fresh input arrays each time: LibreOffice passes ByVal arrays by reference)
    Model SignLines()
    PlanStripSigns ops
    out = ApplyOpsToTexts(SignLines(), ops, flagLog)
    ExpectEq "strip", CStr(out(0)), "1." & vbTab & "A device comprising a sensor and a lens."
    ' insert: map from the document itself, added at the first mention per claim
    Model SignLines()
    Set map = CollectSignMap()
    Check "sign map", map.Count = 3, CStr(map.Count)
    Set ops = New Collection
    PlanInsertSigns ops, map
    out = ApplyOpsToTexts(SignLines(), ops, flagLog)
    ExpectEq "insert", CStr(out(2)), "2." & vbTab & "The device (1) of claim 1, wherein the sensor (10, 11) is optical."
    ' map from another document applied to a stripped claim set
    Model L("1.|A device comprising a sensor and a lens.")
    Set ops = New Collection
    PlanInsertSigns ops, map
    out = ApplyOpsToTexts(L("1.|A device comprising a sensor and a lens."), ops, flagLog)
    ExpectEq "insert from map", CStr(out(0)), "1." & vbTab & "A device (1) comprising a sensor (10, 11) and a lens (12)."
End Sub

Private Function SignLines() As Variant
    SignLines = L("1.|A device (1) comprising a sensor (10; 11) and a lens  (12).", "", "2.|The device of claim 1, wherein the sensor is optical.")
End Function
