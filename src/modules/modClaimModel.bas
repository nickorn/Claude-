Attribute VB_Name = "modClaimModel"
Option Explicit

' Claim model shared by all checks and drafting helpers (pure text, no Word
' objects). BuildClaimModel analyses the visible paragraph texts once; the
' results are kept in the public arrays below (claims are numbered 1..gClaimCount
' by position, paragraphs 0..gParaCount-1).

' Reference record in gClaimRefs(c): Array(para, start, length, label, isRangeEnd, kwStart, kwLen, target)
Public Const R_PARA As Long = 0
Public Const R_START As Long = 1
Public Const R_LEN As Long = 2
Public Const R_LABEL As Long = 3
Public Const R_RANGE_END As Long = 4
Public Const R_KW_START As Long = 5
Public Const R_KW_LEN As Long = 6
Public Const R_TARGET As Long = 7     ' claim number, 0 = unknown label

Public gParaCount As Long
Public gTexts() As String
Public gIndents() As Single
Public gLanguage As String            ' "en" or "de"

Public gClaimCount As Long
Public gClaimLabel() As String
Public gClaimPara() As Long           ' first paragraph (with the label)
Public gClaimLastPara() As Long       ' last non-empty paragraph of the claim
Public gClaimBodyStart() As Long      ' position after label + TAB in the first paragraph
Public gClaimPreambleEnd() As Long    ' position in the first paragraph where the preamble ends
Public gClaimCategory() As String     ' method / device / system / program / medium / use / ""
Public gClaimRefs() As Variant        ' Collection per claim: all references (R_*)
Public gClaimParents() As Variant     ' Collection per claim: distinct claim numbers referenced in the preamble
Public gClaimOfPara() As Long         ' claim number per paragraph, 0 = before the first claim

Private mIndex As Collection          ' label key -> first claim number with that label

' Builds the model. texts: visible paragraph texts (0-based); indents: left
' indents in points or Empty; languageSetting: "auto", "en" or "de".
' Returns False if the document contains no claims.
Public Function BuildClaimModel(ByVal texts As Variant, ByVal indents As Variant, _
        ByVal languageSetting As String) As Boolean
    Dim i As Long, c As Long, lbl As String, bodyStart As Long, lo As Long

    lo = LBound(texts)
    gParaCount = UBound(texts) - lo + 1
    gClaimCount = 0
    If gParaCount <= 0 Then Exit Function
    ReDim gTexts(0 To gParaCount - 1)
    ReDim gIndents(0 To gParaCount - 1)
    ReDim gClaimOfPara(0 To gParaCount - 1)
    For i = 0 To gParaCount - 1
        gTexts(i) = CStr(texts(i + lo))
        If Not IsEmpty(indents) Then gIndents(i) = indents(i + LBound(indents))
    Next i

    ' claims and their paragraph extent
    Set mIndex = New Collection
    ReDim gClaimLabel(1 To gParaCount)
    ReDim gClaimPara(1 To gParaCount)
    ReDim gClaimBodyStart(1 To gParaCount)
    For i = 0 To gParaCount - 1
        If ParseClaimLabel(gTexts(i), lbl, bodyStart) Then
            c = c + 1
            gClaimLabel(c) = lbl
            gClaimPara(c) = i
            gClaimBodyStart(c) = bodyStart
            If ClaimNumberOfLabel(lbl) = 0 Then mIndex.Add c, LabelKey(lbl)
        End If
        gClaimOfPara(i) = c
    Next i
    gClaimCount = c
    If c = 0 Then Exit Function

    ReDim Preserve gClaimLabel(1 To c)
    ReDim Preserve gClaimPara(1 To c)
    ReDim Preserve gClaimBodyStart(1 To c)
    ReDim gClaimLastPara(1 To c)
    ReDim gClaimPreambleEnd(1 To c)
    ReDim gClaimCategory(1 To c)
    ReDim gClaimRefs(1 To c)
    ReDim gClaimParents(1 To c)

    For c = 1 To gClaimCount
        gClaimLastPara(c) = gClaimPara(c)
        For i = gClaimPara(c) + 1 To ClaimEndPara(c)
            If Len(Trim$(gTexts(i))) > 0 Then gClaimLastPara(c) = i
        Next i
    Next c

    gLanguage = LCase$(languageSetting)
    If gLanguage <> "en" And gLanguage <> "de" Then gLanguage = DetectLanguage()

    For c = 1 To gClaimCount
        gClaimPreambleEnd(c) = FindPreambleEnd(gTexts(gClaimPara(c)), gClaimBodyStart(c))
        gClaimCategory(c) = DetectCategory(Mid$(gTexts(gClaimPara(c)), gClaimBodyStart(c), _
            gClaimPreambleEnd(c) - gClaimBodyStart(c)))
        CollectClaimReferences c
    Next c
    BuildClaimModel = True
End Function

' Last paragraph index belonging to claim c (up to the next claim, empties included).
Public Function ClaimEndPara(ByVal c As Long) As Long
    If c < gClaimCount Then
        ClaimEndPara = gClaimPara(c + 1) - 1
    Else
        ClaimEndPara = gParaCount - 1
    End If
End Function

' First claim number carrying this label, 0 if none.
Public Function ClaimNumberOfLabel(ByVal label As String) As Long
    On Error Resume Next
    ClaimNumberOfLabel = mIndex.Item(LabelKey(label))
    If Err.Number <> 0 Then ClaimNumberOfLabel = 0
    Err.Clear
End Function

' Text of paragraph p as seen by the checks: for a claim's first paragraph the
' label and TAB are excluded (returned start position tells where the text begins).
Public Function ParaBodyStart(ByVal p As Long) As Long
    Dim c As Long
    c = gClaimOfPara(p)
    ParaBodyStart = 1
    If c > 0 Then
        If gClaimPara(c) = p Then ParaBodyStart = gClaimBodyStart(c)
    End If
End Function

' True if claim c refers to another claim in its preamble ("The method of claim 1, ...").
Public Function IsDependentClaim(ByVal c As Long) As Boolean
    IsDependentClaim = (gClaimParents(c).Count > 0)
End Function

' ---------------------------------------------------------------- analysis

' Counts typical function words in all claim paragraphs.
Private Function DetectLanguage() As String
    Dim p As Long, s As String, en As Long, de As Long
    For p = gClaimPara(1) To gParaCount - 1
        s = " " & LCase$(Replace(gTexts(p), vbTab, " ")) & " "
        en = en + CountOf(s, " the ") + CountOf(s, " a ") + CountOf(s, " an ") + CountOf(s, " of ") + _
            CountOf(s, " and ") + CountOf(s, " claim") + CountOf(s, " wherein ") + CountOf(s, " comprising")
        de = de + CountOf(s, " der ") + CountOf(s, " die ") + CountOf(s, " das ") + CountOf(s, " des ") + _
            CountOf(s, " ein") + CountOf(s, " und ") + CountOf(s, " mit ") + CountOf(s, " nach ") + _
            CountOf(s, "anspruch") + CountOf(s, "anspr" & ChrW(252) & "ch") + CountOf(s, " wobei ")
    Next p
    If de > en Then DetectLanguage = "de" Else DetectLanguage = "en"
End Function

Public Function CountOf(ByVal s As String, ByVal what As String) As Long
    Dim p As Long
    p = InStr(1, s, what)
    Do While p > 0
        CountOf = CountOf + 1
        p = InStr(p + Len(what), s, what)
    Loop
End Function

' The preamble ends at the first ":" or "," or at "comprising"/"wherein"/
' "characterised"/"umfassend"/"aufweisend"/"wobei"/"dadurch gekennzeichnet".
Private Function FindPreambleEnd(ByVal paraText As String, ByVal bodyStart As Long) As Long
    Dim lt As String, best As Long, i As Long, q As Long, marks As Variant
    lt = LCase$(paraText)
    best = Len(paraText) + 1
    marks = Array(":", ",", " comprising", " comprises", " including", " consisting", " wherein", _
        " characterised", " characterized", " umfassend", " aufweisend", " wobei", " dadurch gekennzeichnet", " mit:")
    For i = LBound(marks) To UBound(marks)
        q = InStr(bodyStart, lt, marks(i))
        If q > 0 And q < best Then best = q
    Next i
    FindPreambleEnd = best
End Function

' Category of a claim from the first category word in its preamble.
Public Function DetectCategory(ByVal preamble As String) As String
    Dim words As Variant, i As Long, cat As String, n As Long
    words = SplitWords(preamble)
    For i = LBound(words) To UBound(words)
        cat = CategoryOfWord(CStr(words(i)))
        If Len(cat) > 0 Then
            DetectCategory = cat
            Exit Function
        End If
        n = n + 1
        If n >= 12 Then Exit Function
    Next i
End Function

Public Function CategoryOfWord(ByVal w As String) As String
    w = LCase$(w)
    If Len(w) < 3 Then Exit Function
    Select Case True
        Case w = "method" Or w = "methods" Or w = "process" Or w = "procedure" Or EndsWith(w, "verfahren") _
                Or EndsWith(w, "verfahrens") Or w = "prozess" Or w = "prozesses"
            CategoryOfWord = "method"
        Case w = "use" Or w = "verwendung"
            CategoryOfWord = "use"
        Case w = "program" Or w = "programs" Or w = "programme" Or InStr(w, "programm") > 0 Or w = "software"
            CategoryOfWord = "program"
        Case w = "medium" Or w = "media" Or EndsWith(w, "medium") Or EndsWith(w, "mediums") Or EndsWith(w, "readable") _
                Or InStr(w, "datentr" & ChrW(228) & "ger") > 0
            CategoryOfWord = "medium"
        Case w = "system" Or w = "systems" Or EndsWith(w, "system") Or EndsWith(w, "systems") Or w = "anlage"
            CategoryOfWord = "system"
        Case w = "device" Or w = "devices" Or w = "apparatus" Or w = "arrangement" Or w = "machine" _
                Or EndsWith(w, "vorrichtung") Or EndsWith(w, "ger" & ChrW(228) & "t") Or EndsWith(w, "ger" & ChrW(228) & "ts") _
                Or EndsWith(w, "einrichtung") Or EndsWith(w, "anordnung") Or w = "apparat"
            CategoryOfWord = "device"
    End Select
End Function

Public Function EndsWith(ByVal s As String, ByVal suffix As String) As Boolean
    If Len(suffix) > Len(s) Then Exit Function
    EndsWith = (Right$(s, Len(suffix)) = suffix)
End Function

' Splits into words of letters/digits/hyphens (everything else separates).
Public Function SplitWords(ByVal s As String) As Variant
    Dim i As Long, c As String, cur As String, out As String
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        If IsLetterChar(c) Or IsDigitChar(c) Or c = "-" Then
            cur = cur & c
        ElseIf Len(cur) > 0 Then
            out = out & Chr$(1) & cur
            cur = ""
        End If
    Next i
    If Len(cur) > 0 Then out = out & Chr$(1) & cur
    If Len(out) = 0 Then
        SplitWords = Array()
    Else
        SplitWords = Split(Mid$(out, 2), Chr$(1))
    End If
End Function

Private Sub CollectClaimReferences(ByVal c As Long)
    Dim p As Long, t As Variant, toks As Collection, target As Long, prevTarget As Long, k As Long
    Dim refs As New Collection, parents As New Collection
    For p = gClaimPara(c) To gClaimLastPara(c)
        Set toks = FindReferences(gTexts(p), ParaBodyStart(p))
        prevTarget = 0
        For Each t In toks
            target = ClaimNumberOfLabel(CStr(t(TK_LABEL)))
            refs.Add Array(p, t(TK_START), t(TK_LEN), t(TK_LABEL), t(TK_RANGE_END), t(TK_KW_START), t(TK_KW_LEN), target)
            If p = gClaimPara(c) And t(TK_KW_START) < gClaimPreambleEnd(c) And target > 0 Then
                If t(TK_RANGE_END) And prevTarget > 0 Then
                    For k = IIf(prevTarget < target, prevTarget, target) To IIf(prevTarget < target, target, prevTarget)
                        AddDistinct parents, k
                    Next k
                Else
                    AddDistinct parents, target
                End If
            End If
            prevTarget = target
        Next t
    Next p
    Set gClaimRefs(c) = refs
    Set gClaimParents(c) = parents
End Sub

Public Sub AddDistinct(ByVal col As Collection, ByVal value As Long)
    Dim v As Variant
    For Each v In col
        If v = value Then Exit Sub
    Next v
    col.Add value
End Sub
