Attribute VB_Name = "modTestTextBlocks"
Option Explicit

' Unit tests for the text-block logic (modTextBlocksPlan).
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunTextBlockTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunTextBlockTests()
    MsgBox RunTextBlockTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunTextBlockTests() As String
    Dim ph As Collection, x As String, ue As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    ue = ChrW(220)

    Check "label prefix", TextBlockLabel("10_Einleitung Art. 94(3) EP" & ue & ".docx") = "Einleitung Art. 94(3) EP" & ue
    Check "label no prefix", TextBlockLabel("Allgemeiner Baustein.docx") = "Allgemeiner Baustein"
    Check "label space prefix", TextBlockLabel("20 Antrag.docx") = "Antrag"
    Check "label digits only name", TextBlockLabel("2024.docx") = "2024"
    Check "label keeps inner digits", TextBlockLabel("30_Art 19 PCT.docx") = "Art 19 PCT"
    Check "file docx", IsTextBlockFile("10_a.docx")
    Check "file temp skipped", Not IsTextBlockFile("~$10_a.docx")
    Check "file pdf skipped", Not IsTextBlockFile("a.pdf")
    Check "xml escape", XmlEscape("A & B <""x""> 'y'") = "A &amp; B &lt;&quot;x&quot;&gt; &apos;y&apos;"

    Set ph = FindPlaceholders("Bescheid vom [DATUM] zu [D1] und [D1], Absatz [0012], [MERKMAL_2], [ ], [x y]")
    Check "placeholders count", ph.Count = 3, CStr(ph.Count)
    If ph.Count = 3 Then Check "placeholders order", ph(1) = "DATUM" And ph(2) = "D1" And ph(3) = "MERKMAL_2"
    Check "placeholder default date", Len(PlaceholderDefault("DATUM")) = 10
    Check "placeholder default other", PlaceholderDefault("D1") = ""

    x = SampleMenuXml()
    Check "menu has submenu", InStr(x, "label=""EPA""") > 0
    Check "menu label stripped", InStr(x, "label=""Einleitung""") > 0
    Check "menu tag path", InStr(x, "tag=""C:\TB\EPA\10_Einleitung.docx""") > 0
    Check "menu escaping", InStr(x, "R&amp;D") > 0
    Check "menu top-level first", InStr(x, "Allgemein") < InStr(x, "label=""EPA""")
    Check "menu sorted", InStr(x, "Einleitung") < InStr(x, "Neuheit")
    Check "menu empty", TextBlockMenuXml("C:\TB", "", "") = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui""></menu>"

    RunTextBlockTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

' Also printed by tools/run_vba_tests.py, which checks that it is well-formed XML.
Public Function SampleMenuXml() As String
    SampleMenuXml = TextBlockMenuXml("C:\TB", "EPA" & vbTab & "50_Neuheit.docx" & vbLf & "EPA" & vbTab & "10_Einleitung.docx" & vbLf & _
        vbTab & "Allgemein R&D.docx" & vbLf & "DPMA" & vbTab & "10_Eingabe.docx", "<button id=""tbOpenFolder"" label=""x""/>")
End Function

Private Sub Check(ByVal name As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If cond Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & name & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
    End If
End Sub
