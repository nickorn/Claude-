Attribute VB_Name = "modConvertPlan"
Option Explicit

' Roadmap 7 drafting helpers, pure part:
'   method <-> apparatus claim, claims -> description paragraphs,
'   reference signs strip / insert.
' The results are drafts for the attorney to review (heuristic wording).

' ---------------------------------------------------------------- English verb forms

Private Function GerundExceptions() As String
    GerundExceptions = "|controlling=control|outputting=output|transmitting=transmit|setting=set|getting=get|" & _
        "putting=put|running=run|storing=store|using=use|providing=provide|receiving=receive|comparing=compare|" & _
        "combining=combine|determining=determine|generating=generate|computing=compute|encoding=encode|" & _
        "decoding=decode|measuring=measure|authorizing=authorize|authorising=authorise|executing=execute|" & _
        "causing=cause|processing=process|obtaining=obtain|sending=send|performing=perform|filtering=filter|" & _
        "detecting=detect|displaying=display|emitting=emit|submitting=submit|including=include|" & _
        "creating=create|updating=update|deleting=delete|retrieving=retrieve|identifying=identify|"
End Function

Private Function KeepCase(ByVal original As String, ByVal w As String) As String
    If Len(original) > 0 And Len(w) > 0 Then
        If Left$(original, 1) = UCase$(Left$(original, 1)) And Left$(original, 1) <> LCase$(Left$(original, 1)) Then
            w = UCase$(Left$(w, 1)) & Mid$(w, 2)
        End If
    End If
    KeepCase = w
End Function

' "obtaining" -> "obtain", "receiving" -> "receive", "outputting" -> "output"
Public Function GerundToVerb(ByVal w As String) As String
    Dim lw As String, q As Long, stem As String, ex As String
    lw = LCase$(w)
    GerundToVerb = w
    If Len(lw) < 6 Or Not EndsWith(lw, "ing") Then Exit Function
    ex = GerundExceptions()
    q = InStr(1, ex, "|" & lw & "=")
    If q > 0 Then
        stem = Mid$(ex, q + Len(lw) + 2, InStr(q + 1, ex, "|") - q - Len(lw) - 2)
    Else
        stem = Left$(lw, Len(lw) - 3)
        If Len(stem) >= 3 And Right$(stem, 1) = Mid$(stem, Len(stem) - 1, 1) And InStr("lsfzaeiou", Right$(stem, 1)) = 0 Then
            stem = Left$(stem, Len(stem) - 1)
        ElseIf EndsWithAny(stem, "|at|iz|is|ys|yz|ur|ar|in|od|id|ut|us|ov|iv|ag|uc|rv|rc|lv|tl|bl|pl|") Then
            stem = stem & "e"
        End If
    End If
    GerundToVerb = KeepCase(w, stem)
End Function

' "receive" -> "receiving", "transmit" -> "transmitting", "obtain" -> "obtaining"
Public Function VerbToGerund(ByVal w As String) As String
    Dim lw As String, ex As String, q As Long, g As String
    lw = LCase$(w)
    ex = GerundExceptions()
    q = InStr(1, ex, "=" & lw & "|")
    If q > 0 Then
        g = Mid$(ex, InStrRev(ex, "|", q) + 1, q - InStrRev(ex, "|", q) - 1)
    ElseIf EndsWith(lw, "ie") Then
        g = Left$(lw, Len(lw) - 2) & "ying"
    ElseIf EndsWith(lw, "e") And Not EndsWith(lw, "ee") And Not EndsWith(lw, "ye") And Len(lw) > 2 Then
        g = Left$(lw, Len(lw) - 1) & "ing"
    Else
        g = lw & "ing"
    End If
    VerbToGerund = KeepCase(w, g)
End Function

Private Function EndsWithAny(ByVal s As String, ByVal list As String) As Boolean
    Dim parts As Variant, i As Long
    parts = Split(Mid$(list, 2, Len(list) - 2), "|")
    For i = LBound(parts) To UBound(parts)
        If EndsWith(s, parts(i)) Then
            EndsWithAny = True
            Exit Function
        End If
    Next i
End Function

' Replaces the first word of t using fn ("verb" or "gerund"); keeps leading text.
Private Function ConvertFirstWord(ByVal t As String, ByVal toVerb As Boolean) As String
    Dim i As Long, e As Long, w As String
    i = 1
    Do While i <= Len(t)
        If IsLetterChar(Mid$(t, i, 1)) Then Exit Do
        i = i + 1
    Loop
    e = i
    Do While e <= Len(t)
        If Not IsLetterChar(Mid$(t, e, 1)) Then Exit Do
        e = e + 1
    Loop
    If e = i Then
        ConvertFirstWord = t
        Exit Function
    End If
    w = Mid$(t, i, e - i)
    If toVerb Then w = GerundToVerb(w) Else w = VerbToGerund(w)
    ConvertFirstWord = Left$(t, i - 1) & w & Mid$(t, e)
End Function

' Case-insensitive replacement of every occurrence (portable; avoids the
' optional compare argument of Replace).
Public Function ReplaceText(ByVal t As String, ByVal findText As String, ByVal repl As String) As String
    Dim q As Long
    If Len(findText) = 0 Then
        ReplaceText = t
        Exit Function
    End If
    q = InStr(1, LCase$(t), LCase$(findText))
    Do While q > 0
        t = Left$(t, q - 1) & repl & Mid$(t, q + Len(findText))
        q = InStr(q + Len(repl), LCase$(t), LCase$(findText))
    Loop
    ReplaceText = t
End Function

' Case-insensitive whole-word replacement.
Public Function ReplaceWord(ByVal t As String, ByVal word As String, ByVal repl As String) As String
    Dim lt As String, lw As String, q As Long, okB As Boolean, okA As Boolean
    lt = LCase$(t)
    lw = LCase$(word)
    q = InStr(1, lt, lw)
    Do While q > 0
        okB = True
        If q > 1 Then okB = Not IsLetterChar(Mid$(lt, q - 1, 1))
        okA = Not IsLetterChar(Mid$(lt, q + Len(lw), 1))
        If okB And okA Then
            t = Left$(t, q - 1) & KeepCase(Mid$(t, q, Len(word)), repl) & Mid$(t, q + Len(word))
            lt = LCase$(t)
            q = InStr(q + Len(repl), lt, lw)
        Else
            q = InStr(q + 1, lt, lw)
        End If
    Loop
    ReplaceWord = t
End Function

' ---------------------------------------------------------------- method <-> apparatus

' Returns the counterpart of claim c as paragraph texts without label
' (Array(ok, message, paragraphs()) - paragraphs as a 0-based String array).
Public Function ConvertClaimCategory(ByVal c As Long) As Variant
    Dim cat As String, noun As String, first As String, pre As String, rest As String
    Dim out() As String, n As Long, p As Long, t As String, k As Long, toMethod As Boolean, listOfVerbs As Boolean
    Dim marks As Variant, i As Long

    cat = gClaimCategory(c)
    If IsDependentClaim(c) Then
        ConvertClaimCategory = Array(False, MsgConvertDependent(), Empty)
        Exit Function
    End If
    If cat <> "method" And cat <> "device" And cat <> "system" Then
        ConvertClaimCategory = Array(False, MsgConvertCategory(CategoryName(cat)), Empty)
        Exit Function
    End If
    toMethod = (cat <> "method")
    noun = CategoryNoun(c)
    first = gTexts(gClaimPara(c))
    pre = Mid$(first, gClaimBodyStart(c), gClaimPreambleEnd(c) - gClaimBodyStart(c))
    rest = Mid$(first, gClaimPreambleEnd(c))
    ReDim out(0 To gClaimLastPara(c) - gClaimPara(c))

    If gLanguage = "de" Then
        If toMethod Then
            pre = ReplaceWord(pre, noun, "Verfahren")
            rest = ReplaceDePhrases(rest, Array(", die eingerichtet ist zum:", ", das eingerichtet ist zum:", _
                ", eingerichtet zum:", ", umfassend:", ", aufweisend:"), ", mit:")
        Else
            pre = ReplaceWord(pre, noun, "Vorrichtung")
            rest = ReplaceDePhrases(rest, Array(", das Verfahren umfassend:", ", umfassend:", ", aufweisend:", _
                ", mit:", ", mit den Schritten:"), ", die eingerichtet ist zum:")
        End If
    Else
        If toMethod Then
            ' "An apparatus configured to:" has the marker in the preamble itself
            pre = SwapNounWithArticle(pre, noun, "method")
            rest = ReplaceWord(rest, noun, "method")
            marks = Array(" being configured to", " configured to", " adapted to", " arranged to", " operable to")
            For i = LBound(marks) To UBound(marks)
                If InStr(1, LCase$(pre & rest), marks(i)) > 0 Then
                    pre = ReplaceText(pre, CStr(marks(i)), ", comprising")
                    rest = ReplaceText(rest, CStr(marks(i)), " comprising")
                    listOfVerbs = True
                End If
            Next i
        Else
            pre = SwapNounWithArticle(pre, noun, "apparatus")
            rest = ReplaceText(rest, "the " & noun & " further comprising", "the apparatus being further configured to")
            rest = ReplaceText(rest, "the " & noun & " comprising", "the apparatus being configured to")
            rest = ReplaceText(rest, "comprising", "configured to")
        End If
    End If
    out(0) = pre & rest

    n = 0
    For p = gClaimPara(c) + 1 To gClaimLastPara(c)
        t = gTexts(p)
        If Not IsBlankText(t) Then
            If gLanguage = "de" Then
                If toMethod Then
                    k = InStr(1, t, " zum ")
                    If k = 0 Then k = InStr(1, t, " zur ")
                    If k > 0 Then t = UCase$(Mid$(t, k + 5, 1)) & Mid$(t, k + 6)
                End If
            ElseIf toMethod Then
                k = ConfiguredToEnd(t)
                If k > 0 Then
                    t = ConvertFirstWord(Mid$(t, k), False)
                ElseIf listOfVerbs Then
                    t = ConvertFirstWord(t, False)
                End If
                t = ReplaceWord(t, noun, "method")
            Else
                t = ReplaceText(t, "the " & noun & " further comprising", "the apparatus being further configured to")
                t = ConvertFirstWord(t, True)
                t = ReplaceWord(t, noun, "apparatus")
            End If
        End If
        n = n + 1
        out(n) = t
    Next p
    ConvertClaimCategory = Array(True, "", out)
End Function

Private Function ReplaceDePhrases(ByVal t As String, ByVal froms As Variant, ByVal repl As String) As String
    Dim i As Long
    For i = LBound(froms) To UBound(froms)
        If InStr(1, LCase$(t), LCase$(froms(i))) > 0 Then
            ReplaceDePhrases = ReplaceText(t, CStr(froms(i)), repl)
            Exit Function
        End If
    Next i
    ReplaceDePhrases = t
End Function

' "A method for x" -> "An apparatus for x"; "Method for x" -> "Apparatus for x"
Private Function SwapNounWithArticle(ByVal pre As String, ByVal noun As String, ByVal newNoun As String) As String
    Dim t As String, art As String
    t = ReplaceWord(pre, noun, newNoun)
    If LCase$(Left$(t, 2)) = "a " Or LCase$(Left$(t, 3)) = "an " Then
        If InStr("aeiou", LCase$(Left$(newNoun, 1))) > 0 Then art = "An " Else art = "A "
        t = art & Mid$(t, InStr(t, " ") + 1)
    End If
    SwapNounWithArticle = t
End Function

' Position after "configured to " etc. in a feature, 0 if none.
Private Function ConfiguredToEnd(ByVal t As String) As Long
    Dim marks As Variant, i As Long, q As Long
    marks = Array(" configured to ", " adapted to ", " arranged to ", " operable to ")
    For i = LBound(marks) To UBound(marks)
        q = InStr(1, LCase$(t), marks(i))
        If q > 0 Then
            ConfiguredToEnd = q + Len(marks(i))
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------- claims -> description

' One description paragraph for claim c, starting with intro (e.g.
' "According to an embodiment, " / "Gemaess einer Ausfuehrungsform ").
Public Function ClaimToDescription(ByVal c As Long, ByVal intro As String) As String
    Dim pre As String, rest As String, p As Long, noun As String, art As String, lt As String

    pre = Mid$(gTexts(gClaimPara(c)), gClaimBodyStart(c), gClaimPreambleEnd(c) - gClaimBodyStart(c))
    rest = RTrimText(Mid$(gTexts(gClaimPara(c)), gClaimPreambleEnd(c)))
    For p = gClaimPara(c) + 1 To gClaimLastPara(c)
        If Not IsBlankText(gTexts(p)) Then rest = rest & " " & Trim$(RTrimText(gTexts(p)))
    Next p
    noun = CategoryNoun(c)

    If Not IsDependentClaim(c) Then
        If gLanguage = "de" Then
            art = GermanArticle(noun, False)
            If LCase$(Left$(pre, 4)) = "ein " Or LCase$(Left$(pre, 5)) = "eine " Then art = ""
            ClaimToDescription = intro & "wird " & IIf(Len(art) > 0, art & " ", "") & LowerArticle(pre) & " bereitgestellt" & rest
        Else
            If Not (LCase$(Left$(pre, 2)) = "a " Or LCase$(Left$(pre, 3)) = "an ") Then
                If InStr("aeiou", LCase$(Left$(pre, 1))) > 0 Then art = "an " Else art = "a "
            End If
            ClaimToDescription = intro & art & LowerFirst(pre) & " is provided" & rest
        End If
        Exit Function
    End If

    ' dependent claim: only what it adds
    Do While Len(rest) > 0
        If Left$(rest, 1) <> "," And Not IsSpaceChar(Left$(rest, 1)) Then Exit Do
        rest = Mid$(rest, 2)
    Loop
    lt = LCase$(rest)
    If gLanguage = "de" Then
        If Left$(lt, 6) = "wobei " Then
            ClaimToDescription = intro & "ist vorgesehen, dass " & Mid$(rest, 7)
        ElseIf Left$(lt, 17) = "ferner umfassend " Or Left$(lt, 18) = "ferner umfassend:" Then
            ClaimToDescription = intro & "umfasst " & GermanArticle(noun, True) & " " & noun & " ferner " & Trim$(Mid$(rest, 18))
        Else
            ClaimToDescription = intro & "ist vorgesehen, dass " & rest
        End If
    Else
        If Left$(lt, 8) = "wherein " Then
            ClaimToDescription = intro & Mid$(rest, 9)
        ElseIf Left$(lt, 19) = "further comprising " Or Left$(lt, 19) = "further comprising:" Then
            ClaimToDescription = intro & "the " & LCase$(noun) & " further comprises " & Trim$(Mid$(rest, 20))
        Else
            ClaimToDescription = intro & rest
        End If
    End If
End Function

Private Function LowerFirst(ByVal s As String) As String
    LowerFirst = LCase$(Left$(s, 1)) & Mid$(s, 2)
End Function

' German: only a leading article is lower-cased ("Eine Vorrichtung" -> "eine Vorrichtung").
Private Function LowerArticle(ByVal s As String) As String
    If Left$(s, 4) = "Ein " Or Left$(s, 5) = "Eine " Then
        LowerArticle = LowerFirst(s)
    Else
        LowerArticle = s
    End If
End Function

' Indefinite (ein/eine) or definite (das/die) article from the noun ending.
Private Function GermanArticle(ByVal noun As String, ByVal definite As Boolean) As String
    Dim ln As String, feminine As Boolean
    ln = LCase$(noun)
    feminine = EndsWithAny(ln, "|ung|heit|keit|ion|schaft|t" & ChrW(228) & "t|ik|anlage|einheit|")
    If definite Then
        If feminine Then GermanArticle = "die" Else GermanArticle = "das"
    Else
        If feminine Then GermanArticle = "eine" Else GermanArticle = "ein"
    End If
End Function

' ---------------------------------------------------------------- reference signs

' Deletes every reference-sign group "(150, 450)" in the claims, with the
' spaces in front of it.
Public Sub PlanStripSigns(ByVal ops As Collection)
    Dim p As Long, t As String, q As Long, e As Long, s As Long
    If gClaimCount = 0 Then Exit Sub
    For p = gClaimPara(1) To gParaCount - 1
        t = gTexts(p)
        q = InStr(ParaBodyStart(p), t, "(")
        Do While q > 0
            e = InStr(q + 1, t, ")")
            If e = 0 Then Exit Do
            If IsSignList(Mid$(t, q + 1, e - q - 1)) Then
                s = q
                Do While s > 1
                    If Mid$(t, s - 1, 1) <> " " Then Exit Do
                    s = s - 1
                Loop
                ops.Add Array(OPK_REFERENCE, p, s, e - s + 1, "")
            End If
            q = InStr(e + 1, t, "(")
        Loop
    Next p
End Sub

' Term -> sign map from the claims ("t" & key -> "150, 450" with the
' configured separator); first occurrence wins.
Public Function CollectSignMap() As Collection
    Dim map As New Collection, p As Long, ph As Variant, signs As String, dummy As String
    If gClaimCount = 0 Then
        Set CollectSignMap = map
        Exit Function
    End If
    For p = gClaimPara(1) To gParaCount - 1
        For Each ph In ExtractPhrases(gTexts(p), ParaBodyStart(p))
            If ph(PH_KIND) <> "lead" And Len(ph(PH_SIGNS)) > 0 Then
                signs = Replace(NormaliseSigns(CStr(ph(PH_SIGNS))), ",", gOptRefSignSeparator)
                If Not MapGet(map, "t" & ph(PH_KEY), dummy) Then map.Add signs, "t" & ph(PH_KEY)
            End If
        Next ph
    Next p
    Set CollectSignMap = map
End Function

Public Function MapGet(ByVal map As Collection, ByVal key As String, ByRef value As String) As Boolean
    On Error Resume Next
    value = map.Item(key)
    MapGet = (Err.Number = 0)
    Err.Clear
End Function

' Inserts " (signs)" after the first mention of each mapped term in every
' claim when that mention has no sign yet.
Public Sub PlanInsertSigns(ByVal ops As Collection, ByVal map As Collection)
    Dim c As Long, p As Long, ph As Variant, seen As String, k As String, signs As String
    For c = 1 To gClaimCount
        seen = "|"
        For p = gClaimPara(c) To gClaimLastPara(c)
            For Each ph In ExtractPhrases(gTexts(p), ParaBodyStart(p))
                k = ph(PH_KEY)
                If ph(PH_KIND) <> "lead" And InStr(1, seen, "|" & k & "|") = 0 Then
                    If MapGet(map, "t" & k, signs) Then
                        seen = seen & k & "|"
                        If Len(ph(PH_SIGNS)) = 0 Then ops.Add Array(OPK_REFERENCE, p, ph(PH_END), 0, " (" & signs & ")")
                    End If
                End If
            Next ph
        Next p
    Next c
End Sub
