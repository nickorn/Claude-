Attribute VB_Name = "modTestJson"
Option Explicit

' Unit tests for modJson and the embedded defaults (modDefaults).
' To run in Word: import next to the src modules, Alt+F8 -> PT_RunJsonTests.

Private mLog As String
Private mFailed As Long
Private mPassed As Long

Public Sub PT_RunJsonTests()
    MsgBox RunJsonTests(), vbInformation, "Patent-Tools Tests"
End Sub

Public Function RunJsonTests() As String
    mLog = ""
    mFailed = 0
    mPassed = 0
    TestScalars
    TestNestedLookup
    TestEscapes
    TestEmbeddedDefaults
    TestInvalidJson
    RunJsonTests = "Passed: " & mPassed & "  Failed: " & mFailed & vbLf & mLog
End Function

Private Sub Check(ByVal testName As String, ByVal cond As Boolean, Optional ByVal detail As String = "")
    If cond Then
        mPassed = mPassed + 1
    Else
        mFailed = mFailed + 1
        mLog = mLog & "FAIL " & testName & IIf(Len(detail) > 0, ": " & detail, "") & vbLf
    End If
End Sub

Private Sub TestScalars()
    Dim root As Object
    Set root = JsonParse("{""s"": ""x"", ""n"": 1.5, ""i"": -3, ""t"": true, ""f"": false, ""z"": null, ""a"": [1, 2]}")
    Check "string", JsonLookup(root, "s") = "x"
    Check "number", JsonLookup(root, "n") = 1.5
    Check "negative", JsonLookup(root, "i") = -3
    Check "true", JsonLookup(root, "t") = True
    Check "false", JsonLookup(root, "f") = False
    Check "null", IsNull(JsonLookup(root, "z"))
    Check "array is not a scalar", IsEmpty(JsonLookup(root, "a"))
    Check "missing", IsEmpty(JsonLookup(root, "nope"))
End Sub

Private Sub TestNestedLookup()
    Dim root As Object
    Set root = JsonParse("{ ""a"": { ""b"": { ""c"": ""deep"" } }, ""A"": 1 }")
    Check "nested", JsonLookup(root, "a.b.c") = "deep"
    Check "case-sensitive keys", JsonLookup(root, "A") = 1
    Check "path through scalar", IsEmpty(JsonLookup(root, "a.b.c.d"))
End Sub

Private Sub TestEscapes()
    Dim root As Object
    Set root = JsonParse("{""p"": ""%APPDATA%\\PatentTools"", ""q"": ""say \""hi\"""", ""u"": ""gro\u00df\u00e4"", ""tab"": ""a\tb""}")
    Check "backslash", JsonLookup(root, "p") = "%APPDATA%\PatentTools"
    Check "quote", JsonLookup(root, "q") = "say ""hi"""
    Check "unicode", JsonLookup(root, "u") = "gro" & ChrW(223) & ChrW(228)
    Check "tab", JsonLookup(root, "tab") = "a" & vbTab & "b"
    Check "decode non-ascii only", DecodeJsonUnicode("x\u00fcy\u0022") = "x" & ChrW(252) & "y\u0022"
End Sub

Private Sub TestEmbeddedDefaults()
    Dim root As Object
    Set root = JsonParse(DecodeJsonUnicode(DefaultSettingsJson()))
    Check "default trackChangesMode", JsonLookup(root, "renumber.trackChangesMode") = "tracked"
    Check "default comment author", JsonLookup(root, "check.commentAuthor") = "PatentCheck"
    Check "default font", JsonLookup(root, "format.fontName") = "Arial"
    Check "default textblock folder", JsonLookup(root, "textblocks.folder") = "%APPDATA%\PatentTools\textblocks"
    Check "default footer DE", JsonLookup(root, "footer.status.de.clean") = "Reinschrift"
    Set root = JsonParse(DecodeJsonUnicode(DefaultBlacklistJson()))
    Check "blacklist parses", root.Exists("de")
    Check "blacklist umlaut", root.Item("de").Item(2) = "ungef" & ChrW(228) & "hr"
End Sub

Private Sub TestInvalidJson()
    Dim v As Variant
    On Error Resume Next
    v = JsonParse("{""a"": }")
    Check "invalid raises", Err.Number <> 0
    Err.Clear
    v = JsonParse("{""a"": ""unterminated}")
    Check "unterminated raises", Err.Number <> 0
    Err.Clear
End Sub
