Attribute VB_Name = "modJson"
Option Explicit

' Minimal JSON parser (own code, no third-party dependency, fully offline).
' Objects -> Scripting.Dictionary (case-sensitive keys), arrays -> Collection,
' strings -> String, numbers -> Double, true/false -> Boolean, null -> Null.

Private mText As String
Private mPos As Long

Public Function JsonParse(ByVal jsonText As String) As Variant
    Dim v As Variant
    mText = jsonText
    mPos = 1
    If Len(mText) > 0 Then
        If AscW(Left$(mText, 1)) = &HFEFF Then mPos = 2 ' UTF-8 BOM
    End If
    ParseValue v
    SkipWs
    If mPos <= Len(mText) Then Fail "unexpected characters after the end"
    If IsObject(v) Then
        Set JsonParse = v
    Else
        JsonParse = v
    End If
End Function

' Looks up a scalar by dotted path ("renumber.trackChangesMode").
' Returns Empty if the path does not exist or does not end in a scalar.
Public Function JsonLookup(ByVal root As Object, ByVal path As String) As Variant
    Dim parts As Variant, i As Long, cur As Object
    JsonLookup = Empty
    If root Is Nothing Then Exit Function
    parts = Split(path, ".")
    Set cur = root
    For i = LBound(parts) To UBound(parts)
        If TypeOf cur Is Collection Then Exit Function ' arrays have no named keys
        If Not cur.Exists(parts(i)) Then Exit Function
        If i = UBound(parts) Then
            If Not IsObject(cur.Item(parts(i))) Then JsonLookup = cur.Item(parts(i))
            Exit Function
        End If
        If Not IsObject(cur.Item(parts(i))) Then Exit Function
        Set cur = cur.Item(parts(i))
    Next i
End Function

' Replaces \uXXXX escapes for non-ASCII characters by the character itself.
' Used for the embedded defaults, which are stored ASCII-only (see modDefaults).
Public Function DecodeJsonUnicode(ByVal s As String) As String
    Dim p As Long, code As Long
    p = InStr(1, s, "\u")
    Do While p > 0
        code = HexValue(Mid$(s, p + 2, 4))
        If code >= &H80 Then
            s = Left$(s, p - 1) & ChrW(code) & Mid$(s, p + 6)
        Else
            p = p + 2
        End If
        p = InStr(p, s, "\u")
    Loop
    DecodeJsonUnicode = s
End Function

' ---------------------------------------------------------------- parser

Private Sub Fail(ByVal what As String)
    Err.Raise vbObjectError + 701, "modJson", "JSON: " & what & " (Zeichen " & mPos & ")"
End Sub

Private Sub SkipWs()
    Dim c As String
    Do While mPos <= Len(mText)
        c = Mid$(mText, mPos, 1)
        If c <> " " And c <> vbTab And c <> vbCr And c <> vbLf Then Exit Do
        mPos = mPos + 1
    Loop
End Sub

Private Sub Expect(ByVal literal As String)
    If Mid$(mText, mPos, Len(literal)) <> literal Then Fail "expected " & literal
    mPos = mPos + Len(literal)
End Sub

Private Sub ParseValue(ByRef v As Variant)
    SkipWs
    Select Case Mid$(mText, mPos, 1)
        Case "{"
            Set v = ParseObject()
        Case "["
            Set v = ParseArray()
        Case """"
            v = ParseString()
        Case "t"
            Expect "true"
            v = True
        Case "f"
            Expect "false"
            v = False
        Case "n"
            Expect "null"
            v = Null
        Case ""
            Fail "unexpected end"
        Case Else
            v = ParseNumber()
    End Select
End Sub

Private Function ParseObject() As Object
    Dim d As Object, key As String, v As Variant
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 0 ' binary: keys are case-sensitive
    mPos = mPos + 1
    SkipWs
    If Mid$(mText, mPos, 1) = "}" Then
        mPos = mPos + 1
        Set ParseObject = d
        Exit Function
    End If
    Do
        SkipWs
        If Mid$(mText, mPos, 1) <> """" Then Fail "expected key"
        key = ParseString()
        SkipWs
        Expect ":"
        ParseValue v
        If d.Exists(key) Then d.Remove key
        d.Add key, v
        SkipWs
        Select Case Mid$(mText, mPos, 1)
            Case ","
                mPos = mPos + 1
            Case "}"
                mPos = mPos + 1
                Exit Do
            Case Else
                Fail "expected , or }"
        End Select
    Loop
    Set ParseObject = d
End Function

Private Function ParseArray() As Collection
    Dim c As New Collection, v As Variant
    mPos = mPos + 1
    SkipWs
    If Mid$(mText, mPos, 1) = "]" Then
        mPos = mPos + 1
        Set ParseArray = c
        Exit Function
    End If
    Do
        ParseValue v
        c.Add v
        SkipWs
        Select Case Mid$(mText, mPos, 1)
            Case ","
                mPos = mPos + 1
            Case "]"
                mPos = mPos + 1
                Exit Do
            Case Else
                Fail "expected , or ]"
        End Select
    Loop
    Set ParseArray = c
End Function

Private Function ParseString() As String
    Dim sb As String, q As Long, b As Long
    mPos = mPos + 1 ' opening quote
    Do
        q = InStr(mPos, mText, """")
        If q = 0 Then Fail "unterminated string"
        b = InStr(mPos, mText, "\")
        If b > 0 And b < q Then
            sb = sb & Mid$(mText, mPos, b - mPos)
            Select Case Mid$(mText, b + 1, 1)
                Case """", "\", "/"
                    sb = sb & Mid$(mText, b + 1, 1)
                Case "b"
                    sb = sb & ChrW(8)
                Case "f"
                    sb = sb & ChrW(12)
                Case "n"
                    sb = sb & vbLf
                Case "r"
                    sb = sb & vbCr
                Case "t"
                    sb = sb & vbTab
                Case "u"
                    sb = sb & ChrW(HexValue(Mid$(mText, b + 2, 4)))
                    b = b + 4
                Case Else
                    mPos = b
                    Fail "invalid escape"
            End Select
            mPos = b + 2
        Else
            sb = sb & Mid$(mText, mPos, q - mPos)
            mPos = q + 1
            Exit Do
        End If
    Loop
    ParseString = sb
End Function

Private Function ParseNumber() As Double
    Dim st As Long, c As String
    st = mPos
    Do While mPos <= Len(mText)
        c = Mid$(mText, mPos, 1)
        If InStr(1, "+-0123456789.eE", c) = 0 Then Exit Do
        mPos = mPos + 1
    Loop
    If mPos = st Then Fail "unexpected character"
    ParseNumber = Val(Mid$(mText, st, mPos - st)) ' Val always uses "." as decimal point
End Function

Private Function HexValue(ByVal h As String) As Long
    Dim i As Long, d As Long
    If Len(h) <> 4 Then Fail "invalid \u escape"
    For i = 1 To 4
        d = InStr(1, "0123456789abcdef", LCase$(Mid$(h, i, 1))) - 1
        If d < 0 Then Fail "invalid \u escape"
        HexValue = HexValue * 16 + d
    Next i
End Function
