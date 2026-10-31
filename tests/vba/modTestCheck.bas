Attribute VB_Name = "modTestCheck"
Option Explicit

' Unit tests for the MVP-2 text checks (modClaimModel, modCheckStructure,
' modCheckTerms). Includes EP 20 187 485.6 from docs/reference with the
' expected findings listed there.
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunCheckTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunCheckTests()
    MsgBox RunCheckTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunCheckTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    TestModel
    TestNumbering
    TestPunctuationEnglish
    TestPunctuationGerman
    TestSpacing
    TestBackReferences
    TestIndependentClaims
    TestMultipleDependencies
    TestCategories
    TestAntecedentEnglish
    TestAntecedentGerman
    TestReferenceSigns
    TestBlacklist
    TestReferenceOriginal
    TestReferenceAmended
    RunCheckTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
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

Private Function Run1(ByVal lines As Variant, ByVal checkKey As String) As Collection
    Dim findings As New Collection
    SetDefaultCheckOptions
    If BuildClaimModel(lines, Empty, "auto") Then RunTextCheck checkKey, findings
    Set Run1 = findings
End Function

Private Function Describe(ByVal findings As Collection) As String
    Dim f As Variant
    For Each f In findings
        Describe = Describe & vbLf & "    [" & f(F_CHECK) & "] para " & f(F_PARA) & " pos " & f(F_START) & ": " & f(F_MSG)
    Next f
End Function

Private Sub ExpectCount(ByVal testName As String, ByVal findings As Collection, ByVal expected As Long)
    Check testName, findings.Count = expected, "got " & findings.Count & " expected " & expected & Describe(findings)
End Sub

Private Function HasMsg(ByVal findings As Collection, ByVal part As String) As Boolean
    Dim f As Variant
    For Each f In findings
        If InStr(1, f(F_MSG), part, vbTextCompare) > 0 Then
            HasMsg = True
            Exit Function
        End If
    Next f
End Function

Private Sub ExpectMsg(ByVal testName As String, ByVal findings As Collection, ByVal part As String)
    Check testName, HasMsg(findings, part), "no finding containing [" & part & "]" & Describe(findings)
End Sub

' ---------------------------------------------------------------- tests

Private Sub TestModel()
    Check "model builds", BuildClaimModel(L("Claims", "1.|A method comprising:", "x;", "y.", "", "2.|The method of claim 1, wherein z."), Empty, "auto")
    Check "model claims", gClaimCount = 2, CStr(gClaimCount)
    Check "model last para", gClaimLastPara(1) = 3, CStr(gClaimLastPara(1))
    Check "model dependent", IsDependentClaim(2) And Not IsDependentClaim(1)
    Check "model category", gClaimCategory(1) = "method" And gClaimCategory(2) = "method", gClaimCategory(1) & "/" & gClaimCategory(2)
    Check "model language en", gLanguage = "en", gLanguage
    Check "model language de", BuildClaimModel(L("1.|Verfahren mit einem Schritt.", "", "2.|Verfahren nach Anspruch 1, wobei der Schritt wiederholt wird."), Empty, "auto") And gLanguage = "de", gLanguage
    Check "category storage medium", DetectCategory("Computer-readable storage medium storing a program") = "medium"
    Check "category program", DetectCategory("Computer program product comprising instructions") = "program"
    Check "category de device", DetectCategory("Steuerungsvorrichtung f" & ChrW(252) & "r ein Fahrzeug") = "device"
    Check "category system", DetectCategory("Authorization system") = "system"
End Sub

Private Sub TestNumbering()
    Dim f As Collection
    Set f = Run1(L("1.|A.", "", "1a.|B.", "", "3.|C.", "", "3.|D."), "numbering")
    ExpectCount "numbering count", f, 2
    ExpectMsg "numbering supplementary", f, "Zusatznummer"
    ExpectMsg "numbering duplicate", f, "mehrfach"
    Set f = Run1(L("1.|A.", "", "3.|B."), "numbering")
    ExpectMsg "numbering gap", f, "erwartet: 2"
    ExpectCount "numbering clean", Run1(L("1.|A.", "", "2.|B."), "numbering"), 0
End Sub

Private Sub TestPunctuationEnglish()
    Dim f As Collection
    ExpectCount "punct ok", Run1(L("1.|A method comprising:", "receiving a signal;", "filtering the signal; and", "outputting the signal."), "punctuation"), 0
    Set f = Run1(L("1.|A method comprising:", "receiving a signal;", "filtering the signal;", "outputting the signal"), "punctuation")
    ExpectCount "punct two errors", f, 2
    ExpectMsg "punct missing and", f, "; and"
    ExpectMsg "punct no full stop", f, "nicht mit einem Punkt"
    Set f = Run1(L("1.|A method comprising:", "receiving a signal", "outputting the signal.."), "punctuation")
    ExpectMsg "punct feature end", f, "Merkmal endet nicht"
    ExpectMsg "punct double stop", f, "mehr als einem Punkt"
    ' "; and" is right before the last item of a nested group (a smaller indent follows) ...
    Dim g As New Collection
    SetDefaultCheckOptions
    BuildClaimModel L("1.|A method comprising:", "a first group:", "x;", "y; and", "z;", "a second step; and", "a last step."), _
        Array(0, 0, 36, 36, 36, 0, 0), "auto"
    CheckPunctuation g
    ExpectCount "punct nested and ok", g, 0
    ' ... and wrong before any other item of the group
    Set g = New Collection
    BuildClaimModel L("1.|A method comprising:", "a first group:", "x; and", "y;", "z;", "a second step; and", "a last step."), _
        Array(0, 0, 36, 36, 36, 0, 0), "auto"
    CheckPunctuation g
    ExpectCount "punct nested and wrong", g, 1
    ExpectMsg "punct nested and wrong msg", g, "nicht vor dem letzten"
End Sub

Private Sub TestPunctuationGerman()
    Dim f As Collection
    ExpectCount "punct de ok", Run1(L("1.|Verfahren mit:", "Empfangen eines Signals;", "Filtern des Signals; und", "Ausgeben des Signals."), "punctuation"), 0
    Set f = Run1(L("1.|Verfahren mit:", "Empfangen eines Signals;", "Filtern des Signals;", "Ausgeben des Signals."), "punctuation")
    ExpectMsg "punct de und", f, "; und"
End Sub

Private Sub TestSpacing()
    Dim f As Collection
    Set f = Run1(L("1.|A device  with a part.", "2.|The device of claim 1 .", "", "", "3.|The device of claim 1 , wherein x."), "spacing")
    ExpectCount "spacing count", f, 5
    ExpectMsg "spacing double", f, "Doppeltes"
    ExpectMsg "spacing 0 empties", f, "stehen 0"
    ExpectMsg "spacing 2 empties", f, "stehen 2"
    ExpectMsg "spacing before punct", f, "Leerzeichen vor"
End Sub

Private Sub TestBackReferences()
    Dim f As Collection
    Set f = Run1(L("1.|A device.", "", "2.|The device of claim 2.", "", "3.|The device of claim 4.", "", "4.|The device of claim 9."), "backReferences")
    ExpectCount "backref count", f, 3
    ExpectMsg "backref self", f, "sich selbst"
    ExpectMsg "backref forward", f, "Vorw"
    ExpectMsg "backref unknown", f, "nicht vorhandenen"
End Sub

Private Sub TestIndependentClaims()
    Dim f As Collection
    Set f = Run1(L("1.|A method comprising a step.", "", "2.|A method comprising performing the method of claim 1.", "", _
        "3.|A device comprising a processor configured to carry out the method of claim 1.", "", _
        "4.|A system configured to carry out the method of claim 1, comprising a server."), "independentClaims")
    ExpectCount "independent count", f, 1
    ExpectMsg "independent msg", f, "derselben Kategorie"
End Sub

Private Sub TestMultipleDependencies()
    Dim f As Collection
    Set f = Run1(L("1.|A method.", "", "2.|The method of claim 1, wherein x.", "", "3.|The method of claim 1 or 2, wherein y.", "", _
        "4.|The method of any one of claims 1 to 3, wherein z."), "multipleDependencies")
    ExpectCount "multidep count", f, 1
    ExpectMsg "multidep on multidep", f, "unzul"
End Sub

Private Sub TestCategories()
    Dim f As Collection
    Set f = Run1(L("1.|A method comprising a step.", "", "2.|The device of claim 1, wherein x.", "", _
        "3.|A device configured to carry out the method of claim 1.", "", "4.|A method for y."), "categories")
    ExpectCount "categories count", f, 2
    ExpectMsg "categories mismatch", f, "passt nicht"
    ExpectMsg "categories second independent", f, "Weiterer unabh"
End Sub

Private Sub TestAntecedentEnglish()
    Dim f As Collection
    Set f = Run1(L("1.|A device comprising a sensor, wherein the processor is connected to the sensors.", "", _
        "2.|The device of claim 1, wherein the sensor is optical and the same sensor is used.", "", _
        "3.|The device of claim 2, comprising a plurality of lenses, wherein the lens is curved."), "antecedentBasis")
    ExpectCount "antecedent en count", f, 3
    ExpectMsg "antecedent en missing", f, "the processor"
    ExpectMsg "antecedent en plural", f, "Singular/Plural"
    ' steps without article introduce terms; suffix of a phrase is accepted
    ExpectCount "antecedent en steps", Run1(L("1.|A method comprising:", "filtering a noisy signal; and", _
        "storing a result of the filtering and the signal."), "antecedentBasis"), 0
End Sub

Private Sub TestAntecedentGerman()
    Dim f As Collection
    Set f = Run1(L("1.|Vorrichtung mit einem Sensor, wobei der Prozessor mit dem Sensor verbunden ist.", "", _
        "2.|Vorrichtung nach Anspruch 1, wobei die Sensoren des ersten Sensors optisch sind."), "antecedentBasis")
    ExpectCount "antecedent de count", f, 2
    ExpectMsg "antecedent de missing", f, "der Prozessor"
    ExpectMsg "antecedent de missing 2", f, "des ersten Sensors"
    ExpectCount "antecedent de steps", Run1(L("1.|Verfahren mit:", "Filtern eines Signals; und", _
        "Speichern des Signals.", "", "2.|Verfahren nach Anspruch 1, wobei das Filtern digital erfolgt."), "antecedentBasis"), 0
End Sub

Private Sub TestReferenceSigns()
    Dim f As Collection
    Set f = Run1(L("1.|A device (1) comprising a sensor (2) and a lens (2, 3), wherein the sensor (4) is behind the lens (2,3).", "", _
        "2.|The device (1) of claim 1, wherein the sensor is optical."), "referenceSigns")
    ExpectMsg "signs used twice", f, "steht bereits"
    ExpectMsg "signs term differs", f, "anderem Bezugszeichen"
    ExpectMsg "signs separator", f, "trennen"
    ExpectMsg "signs missing", f, "fehlt"
    gOptRefSignStyle = "plain"
    BuildClaimModel L("1.|A device (1; 2)."), Empty, "auto"
    Set f = New Collection
    CheckReferenceSigns f
    ExpectCount "signs plain style off", f, 0
End Sub

Private Sub TestBlacklist()
    Dim f As New Collection
    SetDefaultCheckOptions
    gBlacklist.Add "substantially"
    gBlacklist.Add "e.g."
    BuildClaimModel L("Substantially as claimed", "1.|A device being substantially flat, e.g. round, insubstantially thin."), Empty, "auto"
    CheckBlacklist f
    ExpectCount "blacklist count", f, 2
End Sub

' ---------------------------------------------------------------- EP 20 187 485.6

Public Function EpOriginalClaims() As Variant
    EpOriginalClaims = L("1.|Method for authorizing a third party, the method comprising:", "", _
        "performing, at a first entity:", "", _
        "obtaining (1110) a biometric feature of the third party;", "", _
        "obtaining (1120) a secret value;", "", _
        "combining (1130) the biometric feature and the secret value so as to obtain a combined secret value;", "", _
        "performing, at an access controller:", "", _
        "receiving (1210) the secret value; and", "", _
        "receiving (1220) the combined secret value;", "", _
        "processing (1230) the combined secret value using the secret value so as to obtain a decoded biometric feature;", "", _
        "the method further comprising:", "", _
        "obtaining (1310) a biometric feature to be verified;", "", _
        "comparing (1320) the decoded biometric feature with the biometric feature to be verified so as to obtain a comparison result; and", "", _
        "authorizing (1330) the third party based on the comparison result.", "", _
        "2.|The method of claim 1, wherein comparing the decoded biometric feature with the biometric feature is executed at the access controller (150; 450) .")
End Function

Private Function AmendedClaim() As Variant
    AmendedClaim = L("1.|Method for authorizing a third party, the method comprising:", "", _
        "performing, at a first entity:", "", _
        "obtaining (1110) a biometric feature of the third party;", "", _
        "obtaining (1120) a secret value;", "", _
        "combining (1130) the biometric feature and the secret value so as to obtain a combined secret value;", "", _
        "performing, at an access controller:", "", _
        "receiving (1210) the secret value; and", "", _
        "receiving (1220) the combined secret value;", "", _
        "processing (1230) the combined secret value using the secret value so as to obtain a decoded biometric feature;", "", _
        "the method further comprising at the access controller or a different entity:", "", _
        "obtaining (1310) a biometric feature to be verified;", "", _
        "comparing (1320) the decoded biometric feature with the biometric feature to be verified so as to obtain a comparison result; and", "", _
        "authorizing (1330) the third party based on the comparison result such that the third party is authorized for a permission or access at the access controller on behalf of the first entity.", "", _
        "wherein the biometric feature is transmitted from the access controller (150; 450) to the first entity (100) based on an input signal generated by the third party (250);")
End Function

Private Sub TestReferenceOriginal()
    Dim f As Collection
    ' table rows 1, 4, 5, 6 must be found; 7, 8 must NOT be flagged
    Set f = Run1(EpOriginalClaims(), "punctuation")
    ExpectCount "EP orig punctuation", f, 1
    ExpectMsg "EP orig #1 '; and' mid-list", f, "nicht vor dem letzten"
    Set f = Run1(EpOriginalClaims(), "spacing")
    ExpectCount "EP orig spacing", f, 1
    ExpectMsg "EP orig #4 space before full stop", f, "Leerzeichen vor"
    Set f = Run1(EpOriginalClaims(), "referenceSigns")
    ExpectCount "EP orig signs", f, 2
    ExpectMsg "EP orig #5 missing sign", f, "access controller"
    ExpectMsg "EP orig #6 semicolon", f, ";"
    ExpectCount "EP orig #7/#8 antecedent", Run1(EpOriginalClaims(), "antecedentBasis"), 0
    ExpectCount "EP orig numbering", Run1(EpOriginalClaims(), "numbering"), 0
    ExpectCount "EP orig backrefs", Run1(EpOriginalClaims(), "backReferences"), 0
    ExpectCount "EP orig independent", Run1(EpOriginalClaims(), "independentClaims"), 0
    ExpectCount "EP orig categories", Run1(EpOriginalClaims(), "categories"), 0
End Sub

Private Sub TestReferenceAmended()
    Dim f As Collection
    ' table rows 2, 3 (and 1) must be found; row 9 must NOT be flagged
    Set f = Run1(AmendedClaim(), "punctuation")
    ExpectCount "EP amended punctuation", f, 3
    ExpectMsg "EP amended #2 full stop before wherein", f, "Punkt vor dem Ende"
    ExpectMsg "EP amended #3 ends with ;", f, "nicht mit einem Punkt"
    ExpectCount "EP amended #9 antecedent", Run1(AmendedClaim(), "antecedentBasis"), 0
    Set f = Run1(AmendedClaim(), "referenceSigns")
    ExpectMsg "EP amended #5 missing sign", f, "access controller"
    ExpectMsg "EP amended #6 semicolon", f, ";"
End Sub
