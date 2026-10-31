Attribute VB_Name = "modTextBlocksPlan"
Option Explicit

' Text blocks (roadmap 5), pure part: menu labels, ribbon menu XML and
' placeholders. Naming scheme (confirmed by Niklas): <Subfolder>\NN_Label.docx;
' the numeric prefix only sorts, the label is shown without it.

Private Const CUSTOMUI_NS As String = "http://schemas.microsoft.com/office/2009/07/customui"

' "10_Einleitung Art. 94(3) EPUe.docx" -> "Einleitung Art. 94(3) EPUe"
Public Function TextBlockLabel(ByVal fileName As String) As String
    Dim i As Long, s As String
    s = fileName
    i = InStrRev(s, ".")
    If i > 1 Then s = Left$(s, i - 1)
    i = 1
    Do While i <= Len(s)
        If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Do
        i = i + 1
    Loop
    If i > 1 And i < Len(s) Then
        If Mid$(s, i, 1) = "_" Or Mid$(s, i, 1) = " " Then s = Mid$(s, i + 1)
    End If
    TextBlockLabel = Trim$(s)
End Function

Public Function IsTextBlockFile(ByVal fileName As String) As Boolean
    Dim ext As String
    If Left$(fileName, 2) = "~$" Then Exit Function
    ext = LCase$(Mid$(fileName, InStrRev(fileName, ".") + 1))
    IsTextBlockFile = (ext = "docx" Or ext = "docm" Or ext = "doc" Or ext = "dotx" Or ext = "rtf")
End Function

Public Function XmlEscape(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    s = Replace(s, "'", "&apos;")
    XmlEscape = s
End Function

' Sorts an array of strings in place (case-insensitive).
Public Sub SortStrings(ByRef arr As Variant)
    Dim i As Long, j As Long, cur As String
    If Not IsArray(arr) Then Exit Sub
    For i = LBound(arr) + 1 To UBound(arr)
        cur = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If LCase$(arr(j)) <= LCase$(cur) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = cur
    Next i
End Sub

' Builds the dynamicMenu content.
'   basePath  text-block folder (no trailing backslash)
'   listing   one entry per line: subfolder & TAB & file name (subfolder "" = top level)
'   extraXml  buttons appended at the end (folder / import commands)
Public Function TextBlockMenuXml(ByVal basePath As String, ByVal listing As String, ByVal extraXml As String) As String
    Dim lines As Variant, i As Long, parts As Variant, folder As String, cur As String, x As String
    Dim n As Long, path As String, entries() As String, count As Long, menuOpen As Boolean, sortable As Variant

    x = "<menu xmlns=""" & CUSTOMUI_NS & """>"
    If Len(listing) > 0 Then
        lines = Split(listing, vbLf)
        ' sort by subfolder, then file name: top-level files first
        ReDim entries(0 To UBound(lines))
        For i = 0 To UBound(lines)
            parts = Split(lines(i), vbTab)
            If UBound(parts) = 1 Then
                If Len(parts(0)) = 0 Then
                    entries(count) = Chr$(1) & vbTab & parts(1)
                Else
                    entries(count) = lines(i)
                End If
                count = count + 1
            End If
        Next i
        If count > 0 Then
            ReDim Preserve entries(0 To count - 1)
            sortable = entries
            SortStrings sortable
            cur = Chr$(2)
            For i = 0 To count - 1
                parts = Split(sortable(i), vbTab)
                folder = parts(0)
                If folder = Chr$(1) Then folder = ""
                If folder <> cur Then
                    If menuOpen Then x = x & "</menu>"
                    menuOpen = False
                    If Len(folder) > 0 Then
                        n = n + 1
                        x = x & "<menu id=""tbMenu" & n & """ label=""" & XmlEscape(folder) & """>"
                        menuOpen = True
                    End If
                    cur = folder
                End If
                n = n + 1
                If Len(folder) > 0 Then
                    path = basePath & "\" & folder & "\" & parts(1)
                Else
                    path = basePath & "\" & parts(1)
                End If
                x = x & "<button id=""tbItem" & n & """ label=""" & XmlEscape(TextBlockLabel(CStr(parts(1)))) & _
                    """ tag=""" & XmlEscape(path) & """ onAction=""OnAction_InsertTextBlock""/>"
            Next i
            If menuOpen Then x = x & "</menu>"
        End If
    End If
    If n > 0 And Len(extraXml) > 0 Then x = x & "<menuSeparator id=""tbSep""/>"
    TextBlockMenuXml = x & extraXml & "</menu>"
End Function

' Distinct placeholders "[NAME]" in order of appearance (letters, digits, _ and
' -; at least one letter, so paragraph numbers like [0012] are no placeholders).
Public Function FindPlaceholders(ByVal t As String) As Collection
    Dim found As New Collection, q As Long, e As Long, name As String, seen As String
    q = InStr(1, t, "[")
    Do While q > 0
        e = InStr(q + 1, t, "]")
        If e = 0 Then Exit Do
        name = Mid$(t, q + 1, e - q - 1)
        If IsPlaceholderName(name) Then
            If InStr(1, seen, Chr$(1) & name & Chr$(1)) = 0 Then
                seen = seen & Chr$(1) & name & Chr$(1)
                found.Add name
            End If
        End If
        q = InStr(q + 1, t, "[")
    Loop
    Set FindPlaceholders = found
End Function

Private Function IsPlaceholderName(ByVal name As String) As Boolean
    Dim i As Long, c As String, hasLetter As Boolean
    If Len(name) = 0 Or Len(name) > 30 Then Exit Function
    For i = 1 To Len(name)
        c = Mid$(name, i, 1)
        If IsLetterChar(c) Then
            hasLetter = True
        ElseIf Not (IsDigitChar(c) Or c = "_" Or c = "-") Then
            Exit Function
        End If
    Next i
    IsPlaceholderName = hasLetter
End Function

' Default value offered for a placeholder (today's date for [DATUM]/[DATE]).
Public Function PlaceholderDefault(ByVal name As String) As String
    Select Case UCase$(name)
        Case "DATUM", "DATE"
            PlaceholderDefault = Format$(Date, "dd.mm.yyyy")
    End Select
End Function

' Expands %VARIABLE% parts of a path with the environment.
Public Function ExpandEnvironment(ByVal path As String) As String
    Dim q As Long, e As Long, name As String
    q = InStr(1, path, "%")
    Do While q > 0
        e = InStr(q + 1, path, "%")
        If e = 0 Then Exit Do
        name = Mid$(path, q + 1, e - q - 1)
        If Len(name) > 0 And Len(Environ$(name)) > 0 Then
            path = Left$(path, q - 1) & Environ$(name) & Mid$(path, e + 1)
            q = InStr(q + Len(Environ$(name)), path, "%")
        Else
            q = InStr(e + 1, path, "%")
        End If
    Loop
    ExpandEnvironment = path
End Function
