Attribute VB_Name = "modCheckTerms"
Option Explicit

' Term-based text checks: antecedent basis, reference signs, blacklist.
' Approach agreed with Niklas: "flag it, you decide" - false positives are
' acceptable, ambiguity ("the biometric feature" with several candidates) is
' NOT flagged.
'
' Noun phrases are found after determiners:
'   EN introduce: a, an, at least one, one or more, a plurality of, ...
'   EN refer:     the, said
'   DE introduce: ein/eine/einen/eines/einem/einer, mindestens/zumindest ein..., mehrere, ...
'   DE refer:     der/die/das/den/dem/des, dieser/diese/..., besagte...
' The first phrase of every paragraph without determiner ("obtaining ...",
' "Verfahren zur ...") also counts as introduced (preamble / method steps).
' German words are compared by a crude stem, so inflections still match.

' Phrase record: Array(kind, key, startPos, endPos, display, signContent)
Public Const PH_KIND As Long = 0      ' "intro", "refer", "lead"
Public Const PH_KEY As Long = 1       ' normalised words joined by " "
Public Const PH_START As Long = 2
Public Const PH_END As Long = 3       ' position after the last word
Public Const PH_DISPLAY As Long = 4
Public Const PH_SIGNS As Long = 5     ' text inside a directly following "( ... )", "" if none

Private Const MAX_PHRASE_WORDS As Long = 5

Private mListsReady As Boolean
Private mHardEn As String, mSoftEn As String, mSkipEn As String
Private mHardDe As String, mSoftDe As String, mSkipDe As String
Private mIntroDe As String, mReferDe As String

' ---------------------------------------------------------------- word lists

Private Sub PrepareLists()
    If mListsReady Then Exit Sub
    mHardEn = "|of|for|to|from|with|by|in|on|at|into|onto|and|or|but|nor|wherein|whereby|whereas|which|that|who|whose|" & _
        "is|are|be|been|being|was|were|has|have|had|comprising|comprises|comprise|including|includes|consisting|" & _
        "configured|arranged|adapted|designed|according|based|such|so|as|between|using|via|than|when|if|after|before|" & _
        "under|over|through|within|without|while|respectively|thereof|therein|said|the|a|an|each|per|not|this|these|" & _
        "those|it|its|their|there|here|claim|claims|can|may|must|shall|will|should|"
    mSoftEn = "|further|furthermore|then|thereby|preferably|optionally|also|only|both|all|"
    mSkipEn = "|same|other|latter|former|following|preceding|respective|invention|art|case|basis|form|way|manner|" & _
        "rest|remainder|end|time|order|sense|purpose|extent|help|aid|"
    mHardDe = Txt("|und|oder|sowie|mit|von|vom|zu|zum|zur|f{ue}r|nach|gem{ae}{ss}|durch|bei|beim|in|im|an|am|auf|aus|" & _
        "{ue}ber|unter|zwischen|wobei|welche|welcher|welches|welchen|welchem|ist|sind|wird|werden|wurde|hat|haben|" & _
        "eingerichtet|ausgebildet|konfiguriert|umfassend|aufweisend|umfasst|aufweist|als|wie|dass|so|um|gegen{ue}ber|" & _
        "ohne|innerhalb|au{ss}erhalb|w{ae}hrend|bzw|mittels|nicht|auch|jeweils|der|die|das|den|dem|des|ein|eine|einen|" & _
        "einem|einer|eines|dieser|diese|dieses|diesen|diesem|anspruch|anspruchs|anspr{ue}che|anspr{ue}chen|kann|k{oe}nnen|")
    mSoftDe = Txt("|ferner|weiter|weiterhin|au{ss}erdem|dann|vorzugsweise|insbesondere|zus{ae}tzlich|")
    mSkipDe = Txt("|gleich|selb|jeweilig|folgend|{ue}brig|erfindung|weis|art|fall|hilf|form|")
    mIntroDe = "|ein|eine|einen|eines|einem|einer|mehrere|mehreren|mehrerer|"
    mReferDe = "|der|die|das|den|dem|des|dieser|diese|dieses|diesen|diesem|besagte|besagten|besagter|besagtes|"
    mListsReady = True
End Sub

Private Function InList(ByVal list As String, ByVal w As String) As Boolean
    InList = (InStr(1, list, "|" & w & "|") > 0)
End Function

' Crude German stem: strips up to two inflection endings, keeps >= 4 letters.
Public Function StemDe(ByVal w As String) As String
    Dim pass As Long, i As Long, endings As Variant, e As String
    endings = Array("en", "em", "er", "es", "e", "n", "s")
    For pass = 1 To 2
        For i = LBound(endings) To UBound(endings)
            e = endings(i)
            If Len(w) - Len(e) >= 4 And EndsWith(w, e) Then
                w = Left$(w, Len(w) - Len(e))
                Exit For
            End If
        Next i
    Next pass
    StemDe = w
End Function

Private Function NormWord(ByVal w As String) As String
    If gLanguage = "de" Then
        NormWord = StemDe(LCase$(w))
    Else
        NormWord = LCase$(w)
    End If
End Function

' ---------------------------------------------------------------- tokenizer

' Splits text from startPos into tokens: words (letters/digits/hyphen/apostrophe)
' and single punctuation characters. Returns the count; arrays are 1-based.
Private Function Tokenize(ByVal t As String, ByVal startPos As Long, ByRef tokStart() As Long, _
        ByRef tokText() As String, ByRef tokIsWord() As Boolean) As Long
    Dim i As Long, c As String, n As Long, inWord As Boolean
    ReDim tokStart(1 To Len(t) + 1)
    ReDim tokText(1 To Len(t) + 1)
    ReDim tokIsWord(1 To Len(t) + 1)
    For i = startPos To Len(t)
        c = Mid$(t, i, 1)
        If IsLetterChar(c) Or IsDigitChar(c) Or ((c = "-" Or c = "'") And inWord) Then
            If Not inWord Then
                n = n + 1
                tokStart(n) = i
                tokText(n) = ""
                tokIsWord(n) = True
                inWord = True
            End If
            tokText(n) = tokText(n) & c
        Else
            inWord = False
            If Not IsSpaceChar(c) Then
                n = n + 1
                tokStart(n) = i
                tokText(n) = c
                tokIsWord(n) = False
            End If
        End If
    Next i
    Tokenize = n
End Function

' Determiner at token i: returns its kind ("intro"/"refer"/"") and the number
' of tokens it spans.
Private Function DeterminerAt(tokText() As String, tokIsWord() As Boolean, ByVal n As Long, _
        ByVal i As Long, ByRef span As Long) As String
    Dim w As String, w2 As String, w3 As String
    If Not tokIsWord(i) Then Exit Function
    w = LCase$(tokText(i))
    If i + 1 <= n Then w2 = LCase$(tokText(i + 1))
    If i + 2 <= n Then w3 = LCase$(tokText(i + 2))
    span = 1
    If gLanguage = "de" Then
        If (w = "mindestens" Or w = "zumindest" Or w = "wenigstens") And InList(mIntroDe, w2) Then
            span = 2
            DeterminerAt = "intro"
        ElseIf (w = "eine" Or w = "einer" Or w = "einen") And (w2 = "mehrzahl" Or w2 = "vielzahl" Or w2 = "anzahl") And w3 = "von" Then
            span = 3
            DeterminerAt = "intro"
        ElseIf InList(mIntroDe, w) Then
            DeterminerAt = "intro"
        ElseIf InList(mReferDe, w) Then
            ' ", die ..." / ", der ..." starts a relative clause
            If i > 1 Then
                If tokText(i - 1) = "," Then Exit Function
            End If
            DeterminerAt = "refer"
        End If
    Else
        If w = "at" And w2 = "least" And w3 = "one" Then
            span = 3
            DeterminerAt = "intro"
        ElseIf (w = "one" Or w = "two") And w2 = "or" And w3 = "more" Then
            span = 3
            DeterminerAt = "intro"
        ElseIf w = "a" And (w2 = "plurality" Or w2 = "number" Or w2 = "set" Or w2 = "pair") And w3 = "of" Then
            span = 3
            DeterminerAt = "intro"
        ElseIf w = "a" Or w = "an" Or w = "several" Or w = "multiple" Then
            DeterminerAt = "intro"
        ElseIf w = "the" Or w = "said" Then
            DeterminerAt = "refer"
        End If
    End If
End Function

Private Function IsHardStop(ByVal w As String) As Boolean
    If gLanguage = "de" Then
        IsHardStop = InList(mHardDe, LCase$(w))
    Else
        IsHardStop = InList(mHardEn, LCase$(w))
    End If
End Function

Private Function IsSoftStop(ByVal w As String) As Boolean
    Dim lw As String
    lw = LCase$(w)
    If gLanguage = "de" Then
        IsSoftStop = InList(mSoftDe, lw)
    Else
        IsSoftStop = InList(mSoftEn, lw) Or (Len(lw) > 4 And (EndsWith(lw, "ing") Or EndsWith(lw, "ed")))
    End If
End Function

' Phrase made of the words starting at token i; returns token index after it.
Private Function PhraseFrom(ByVal t As String, tokStart() As Long, tokText() As String, tokIsWord() As Boolean, _
        ByVal n As Long, ByVal i As Long, ByRef key As String, ByRef pEnd As Long, ByRef display As String) As Long
    Dim k As Long, cnt As Long, w As String, nounSeen As Boolean
    key = ""
    k = i
    Do While k <= n And cnt < MAX_PHRASE_WORDS
        If Not tokIsWord(k) Then Exit Do
        w = tokText(k)
        If IsAllDigits(w) Then Exit Do
        If IsHardStop(w) Then Exit Do
        If cnt > 0 And IsSoftStop(w) Then Exit Do
        ' German: the capitalised head noun ends the phrase ("dem Sensor | verbunden")
        If nounSeen Then Exit Do
        If gLanguage = "de" Then nounSeen = (UCase$(Left$(w, 1)) = Left$(w, 1) And IsLetterChar(Left$(w, 1)))
        If Len(key) > 0 Then key = key & " "
        key = key & NormWord(w)
        pEnd = tokStart(k) + Len(w)
        cnt = cnt + 1
        k = k + 1
    Loop
    If cnt > 0 Then display = Mid$(t, tokStart(i), pEnd - tokStart(i))
    PhraseFrom = k
End Function

' Returns all noun phrases of paragraph text t from startPos (PH_* records).
Public Function ExtractPhrases(ByVal t As String, ByVal startPos As Long) As Collection
    Dim phrases As New Collection, n As Long, i As Long, nextI As Long, span As Long
    Dim tokStart() As Long, tokText() As String, tokIsWord() As Boolean
    Dim kind As String, key As String, pEnd As Long, display As String, signs As String, firstWord As Boolean
    PrepareLists
    n = Tokenize(t, startPos, tokStart, tokText, tokIsWord)
    i = 1
    firstWord = True
    Do While i <= n
        kind = DeterminerAt(tokText, tokIsWord, n, i, span)
        If Len(kind) > 0 Then
            nextI = PhraseFrom(t, tokStart, tokText, tokIsWord, n, i + span, key, pEnd, display)
            If Len(key) > 0 Then
                signs = SignContentAfter(t, pEnd)
                ' span and display include the determiner ("the processor")
                display = Mid$(t, tokStart(i), pEnd - tokStart(i))
                phrases.Add Array(kind, key, tokStart(i), pEnd, display, signs)
                i = nextI
            Else
                i = i + span
            End If
        ElseIf firstWord And tokIsWord(i) Then
            nextI = PhraseFrom(t, tokStart, tokText, tokIsWord, n, i, key, pEnd, display)
            If Len(key) > 0 Then
                phrases.Add Array("lead", key, tokStart(i), pEnd, display, SignContentAfter(t, pEnd))
                i = nextI
            Else
                i = i + 1
            End If
        Else
            i = i + 1
        End If
        firstWord = False
    Loop
    Set ExtractPhrases = phrases
End Function

' If a reference-sign group "( ... )" follows at pos (spaces allowed), returns
' its inner text when it looks like signs (digits with optional letters/primes).
Public Function SignContentAfter(ByVal t As String, ByVal pos As Long) As String
    Dim q As Long, e As Long, inner As String
    q = SkipSpaces(t, pos)
    If Mid$(t, q, 1) <> "(" Then Exit Function
    e = InStr(q + 1, t, ")")
    If e = 0 Then Exit Function
    inner = Mid$(t, q + 1, e - q - 1)
    If IsSignList(inner) Then SignContentAfter = inner
End Function

Public Function IsSignList(ByVal inner As String) As Boolean
    Dim parts As Variant, i As Long, s As String
    parts = Split(Replace(inner, ";", ","), ",")
    For i = LBound(parts) To UBound(parts)
        s = Trim$(parts(i))
        If Len(s) = 0 Then Exit Function
        If Not IsDigitChar(Left$(s, 1)) Then Exit Function
        If Len(s) > 8 Then Exit Function
    Next i
    IsSignList = (UBound(parts) >= LBound(parts))
End Function

' Normalised, sorted sign set: "150,450".
Public Function NormaliseSigns(ByVal inner As String) As String
    Dim parts As Variant, i As Long, j As Long, tmp As String
    parts = Split(Replace(inner, ";", ","), ",")
    For i = LBound(parts) To UBound(parts)
        parts(i) = LCase$(Trim$(parts(i)))
    Next i
    For i = LBound(parts) To UBound(parts) - 1
        For j = i + 1 To UBound(parts)
            If parts(j) < parts(i) Then
                tmp = parts(i)
                parts(i) = parts(j)
                parts(j) = tmp
            End If
        Next j
    Next i
    NormaliseSigns = Join(parts, ",")
End Function

' ---------------------------------------------------------------- antecedent basis

Public Sub CheckAntecedentBasis(ByVal findings As Collection)
    Dim c As Long, p As Long, ph As Variant, own() As String, full() As String, avail As String, k As String
    If gClaimCount = 0 Then Exit Sub
    PrepareLists
    ReDim own(1 To gClaimCount)
    ReDim full(1 To gClaimCount)

    ' pass 1: everything each claim introduces (parents may come later in the document)
    For c = 1 To gClaimCount
        full(c) = "|"
        For p = gClaimPara(c) To gClaimLastPara(c)
            For Each ph In ExtractPhrases(gTexts(p), ParaBodyStart(p))
                If ph(PH_KIND) <> "refer" Then RegisterTerm full(c), CStr(ph(PH_KEY))
            Next ph
        Next p
    Next c

    ' pass 2: each "the ..." needs an earlier introduction in the claim or in a parent claim
    For c = 1 To gClaimCount
        own(c) = "|"
        For p = gClaimPara(c) To gClaimLastPara(c)
            For Each ph In ExtractPhrases(gTexts(p), ParaBodyStart(p))
                k = ph(PH_KEY)
                If ph(PH_KIND) = "refer" Then
                    avail = own(c) & AncestorTerms(c, full)
                    If Not IsSkipWord(k) And Not TermKnown(avail, k) Then
                        If KnownInOtherNumber(avail, k) Then
                            AddFinding findings, "antecedentBasis", p, ph(PH_START), ph(PH_END) - ph(PH_START), _
                                MsgAntecedentPlural(CStr(ph(PH_DISPLAY)))
                        Else
                            AddFinding findings, "antecedentBasis", p, ph(PH_START), ph(PH_END) - ph(PH_START), _
                                MsgAntecedentMissing(CStr(ph(PH_DISPLAY)))
                        End If
                        RegisterTerm own(c), k ' flag each term once per claim
                    End If
                Else
                    RegisterTerm own(c), k
                End If
            Next ph
        Next p
    Next c
End Sub

' Terms of all parent claims, recursively (each claim visited once).
Private Function AncestorTerms(ByVal c As Long, own() As String) As String
    Dim visited As String, stack As New Collection, cur As Long, v As Variant, r As String
    visited = "|"
    For Each v In gClaimParents(c)
        stack.Add v
    Next v
    Do While stack.Count > 0
        cur = stack(stack.Count)
        stack.Remove stack.Count
        If InStr(1, visited, "|" & cur & "|") = 0 And cur <> c Then
            visited = visited & cur & "|"
            r = r & own(cur)
            For Each v In gClaimParents(cur)
                stack.Add v
            Next v
        End If
    Loop
    AncestorTerms = r
End Function

' Registers a phrase and all its leading word groups ("biometric feature" also
' registers "biometric").
Private Sub RegisterTerm(ByRef terms As String, ByVal key As String)
    Dim words As Variant, i As Long, part As String
    words = Split(key, " ")
    For i = LBound(words) To UBound(words)
        If Len(part) > 0 Then part = part & " "
        part = part & words(i)
        If InStr(1, terms, "|" & part & "|") = 0 Then terms = terms & part & "|"
    Next i
End Sub

' Known if registered as such, or as the end of a longer registered phrase
' ("the feature" after "a biometric feature").
Private Function TermKnown(ByVal terms As String, ByVal key As String) As Boolean
    TermKnown = (InStr(1, terms, "|" & key & "|") > 0) Or (InStr(1, terms, " " & key & "|") > 0)
End Function

' True if the term is known in singular/plural form instead (sensor/sensors,
' lens/lenses). German stems already merge singular and plural.
Private Function KnownInOtherNumber(ByVal terms As String, ByVal key As String) As Boolean
    If gLanguage = "de" Then Exit Function
    If TermKnown(terms, key & "s") Or TermKnown(terms, key & "es") Then
        KnownInOtherNumber = True
    ElseIf EndsWith(key, "es") And Len(key) > 4 Then
        KnownInOtherNumber = TermKnown(terms, Left$(key, Len(key) - 2)) Or TermKnown(terms, Left$(key, Len(key) - 1))
    ElseIf EndsWith(key, "s") And Len(key) > 3 Then
        KnownInOtherNumber = TermKnown(terms, Left$(key, Len(key) - 1))
    End If
End Function

Private Function IsSkipWord(ByVal key As String) As Boolean
    Dim first As String, parts As Variant
    parts = Split(key, " ")
    first = parts(0)
    If gLanguage = "de" Then
        IsSkipWord = InList(mSkipDe, first)
    Else
        IsSkipWord = InList(mSkipEn, first)
    End If
End Function

' ---------------------------------------------------------------- reference signs

Public Sub CheckReferenceSigns(ByVal findings As Collection)
    Dim p As Long, c As Long, ph As Variant, signs As String, prev As String, s As Variant
    Dim termSigns As New Collection, signTerm As New Collection, seen As String, k As String
    Dim q As Long, e As Long, t As String, inner As String

    If gOptRefSignStyle <> "parentheses" Or gClaimCount = 0 Then Exit Sub
    PrepareLists

    ' 1. format of every sign group, and term <-> sign consistency
    For p = gClaimPara(1) To gParaCount - 1
        If gClaimOfPara(p) > 0 Then
            t = gTexts(p)
            q = InStr(ParaBodyStart(p), t, "(")
            Do While q > 0
                e = InStr(q + 1, t, ")")
                If e = 0 Then Exit Do
                inner = Mid$(t, q + 1, e - q - 1)
                If IsSignList(inner) Then CheckSignFormat findings, p, q, e - q + 1, inner
                q = InStr(e + 1, t, "(")
            Loop
            For Each ph In ExtractPhrases(t, ParaBodyStart(p))
                If ph(PH_KIND) <> "lead" And Len(ph(PH_SIGNS)) > 0 Then
                    k = ph(PH_KEY)
                    signs = NormaliseSigns(CStr(ph(PH_SIGNS)))
                    If ColGet(termSigns, "t" & k, prev) Then
                        If prev <> signs Then
                            AddFinding findings, "referenceSigns", p, ph(PH_START), ph(PH_END) - ph(PH_START), _
                                MsgSignTermDiffers(CStr(ph(PH_DISPLAY)), signs, prev)
                        End If
                    Else
                        termSigns.Add signs, "t" & k
                    End If
                    For Each s In Split(signs, ",")
                        ' value: key & Chr(1) & display text of the first term with this sign
                        If ColGet(signTerm, "s" & s, prev) Then
                            If Left$(prev, InStr(prev, Chr$(1)) - 1) <> k Then
                                AddFinding findings, "referenceSigns", p, ph(PH_START), ph(PH_END) - ph(PH_START), _
                                    MsgSignUsedTwice(CStr(s), Mid$(prev, InStr(prev, Chr$(1)) + 1))
                            End If
                        Else
                            signTerm.Add k & Chr$(1) & ph(PH_DISPLAY), "s" & s
                        End If
                    Next s
                End If
            Next ph
        End If
    Next p

    ' 2. first mention of a signed term in each claim must carry the sign
    For c = 1 To gClaimCount
        seen = "|"
        For p = gClaimPara(c) To gClaimLastPara(c)
            For Each ph In ExtractPhrases(gTexts(p), ParaBodyStart(p))
                k = ph(PH_KEY)
                If ph(PH_KIND) <> "lead" And InStr(1, seen, "|" & k & "|") = 0 Then
                    If ColGet(termSigns, "t" & k, prev) Then
                        seen = seen & k & "|"
                        If Len(ph(PH_SIGNS)) = 0 Then
                            AddFinding findings, "referenceSigns", p, ph(PH_START), ph(PH_END) - ph(PH_START), _
                                MsgSignMissing(CStr(ph(PH_DISPLAY)), prev)
                        End If
                    End If
                End If
            Next ph
        Next p
    Next c
End Sub

Private Sub CheckSignFormat(ByVal findings As Collection, ByVal p As Long, ByVal q As Long, _
        ByVal length As Long, ByVal inner As String)
    If InStr(inner, ";") > 0 And gOptNormaliseSemicolon Then
        AddFinding findings, "referenceSigns", p, q, length, MsgSignSemicolon(gOptRefSignSeparator)
        Exit Sub
    End If
    If InStr(inner, ",") > 0 Then
        If Replace(Replace(inner, ", ", ","), ",", gOptRefSignSeparator) <> inner Then
            AddFinding findings, "referenceSigns", p, q, length, MsgSignSeparator(gOptRefSignSeparator)
        End If
    End If
End Sub

Private Function ColGet(ByVal col As Collection, ByVal key As String, ByRef value As String) As Boolean
    On Error Resume Next
    value = col.Item(key)
    ColGet = (Err.Number = 0)
    Err.Clear
End Function

' ---------------------------------------------------------------- blacklist

Public Sub CheckBlacklist(ByVal findings As Collection)
    Dim p As Long, term As Variant, lt As String, lterm As String, q As Long, okBefore As Boolean, okAfter As Boolean
    If gBlacklist Is Nothing Or gClaimCount = 0 Then Exit Sub
    For p = gClaimPara(1) To gParaCount - 1
        If gClaimOfPara(p) > 0 Then
            lt = LCase$(gTexts(p))
            For Each term In gBlacklist
                lterm = LCase$(CStr(term))
                If Len(lterm) > 0 Then
                    q = InStr(ParaBodyStart(p), lt, lterm)
                    Do While q > 0
                        okBefore = True
                        If q > 1 Then okBefore = Not (IsLetterChar(Mid$(lt, q - 1, 1)) Or IsDigitChar(Mid$(lt, q - 1, 1)))
                        okAfter = True
                        If IsLetterChar(Right$(lterm, 1)) Or IsDigitChar(Right$(lterm, 1)) Then
                            okAfter = Not (IsLetterChar(Mid$(lt, q + Len(lterm), 1)) Or IsDigitChar(Mid$(lt, q + Len(lterm), 1)))
                        End If
                        If okBefore And okAfter Then
                            AddFinding findings, "blacklist", p, q, Len(lterm), MsgBlacklist(Mid$(gTexts(p), q, Len(lterm)))
                        End If
                        q = InStr(q + Len(lterm), lt, lterm)
                    Loop
                End If
            Next term
        End If
    Next p
End Sub
