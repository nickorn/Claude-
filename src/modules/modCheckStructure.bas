Attribute VB_Name = "modCheckStructure"
Option Explicit

' Structural text checks: numbering, punctuation, spacing, back-references,
' independent claims, multiple dependencies, categories.

' ---------------------------------------------------------------- numbering

Public Sub CheckNumbering(ByVal findings As Collection)
    Dim c As Long, lbl As String, p As Long
    For c = 1 To gClaimCount
        lbl = gClaimLabel(c)
        p = gClaimPara(c)
        If ClaimNumberOfLabel(lbl) <> c Then
            AddFinding findings, "numbering", p, 1, Len(lbl), MsgNumDuplicate(lbl)
        ElseIf Not IsAllDigits(lbl) Then
            AddFinding findings, "numbering", p, 1, Len(lbl), MsgNumSupplementary(lbl)
        ElseIf lbl <> CStr(c) Then
            AddFinding findings, "numbering", p, 1, Len(lbl), MsgNumWrong(lbl, c)
        End If
    Next c
End Sub

' ---------------------------------------------------------------- punctuation

Public Sub CheckPunctuation(ByVal findings As Collection)
    Dim c As Long, p As Long, n As Long, k As Long, t As String, conjLen As Long, listEnd As Long
    Dim ps() As Long
    For c = 1 To gClaimCount
        ' non-empty paragraphs of the claim
        ReDim ps(1 To gClaimLastPara(c) - gClaimPara(c) + 1)
        n = 0
        For p = gClaimPara(c) To gClaimLastPara(c)
            If Not IsBlankText(gTexts(p)) Then
                n = n + 1
                ps(n) = p
            End If
        Next p

        ' end of claim: exactly one full stop
        t = RTrimText(gTexts(ps(n)))
        If Right$(t, 1) <> "." Then
            AddFinding findings, "punctuation", ps(n), Len(t), 1, MsgPunctNoFullStop()
        ElseIf Right$(t, 2) = ".." Then
            AddFinding findings, "punctuation", ps(n), Len(t) - 1, 2, MsgPunctDoubleFullStop()
        End If

        ' the feature list ends at the first full stop (text after it is flagged separately)
        listEnd = n
        For k = 1 To n - 1
            If Right$(RTrimText(gTexts(ps(k))), 1) = "." Then
                listEnd = k
                Exit For
            End If
        Next k

        For k = 1 To n - 1
            t = RTrimText(gTexts(ps(k)))
            conjLen = ConjunctionTailLength(t)
            If Right$(t, 1) = "." Then
                If gOptFlagWherein Then
                    AddFinding findings, "punctuation", ps(k), Len(t), 1, MsgPunctFullStopInside()
                End If
            ElseIf conjLen > 0 Then
                If Not NextIsLastOfList(ps, listEnd, k) Then
                    AddFinding findings, "punctuation", ps(k), Len(t) - conjLen + 1, conjLen, MsgPunctAndMisplaced()
                End If
            ElseIf Right$(t, 1) <> ";" And Right$(t, 1) <> ":" Then
                AddFinding findings, "punctuation", ps(k), IIf(Len(t) > 0, Len(t), 1), 1, MsgPunctFeatureEnd()
            End If
        Next k

        ' "; and" before the last feature
        If gOptEnforceAnd And listEnd >= 3 Then
            t = RTrimText(gTexts(ps(listEnd - 1)))
            If Right$(t, 1) = ";" Then
                AddFinding findings, "punctuation", ps(listEnd - 1), Len(t), 1, MsgPunctAndMissing(gLanguage)
            End If
        End If
    Next c
End Sub

' Length of a trailing "; and" / "; or" / "; und" / "; oder" / "; and/or", 0 if none.
Private Function ConjunctionTailLength(ByVal t As String) As Long
    Dim lt As String, words As Variant, i As Long, w As String, k As Long, wholeWord As Boolean
    lt = LCase$(t)
    words = Array("and/or", "und/oder", "oder", "and", "und", "or")
    For i = LBound(words) To UBound(words)
        w = words(i)
        If EndsWith(lt, w) Then
            k = Len(lt) - Len(w)
            wholeWord = True
            If k > 0 Then wholeWord = Not IsLetterChar(Mid$(lt, k, 1))
            If wholeWord Then
                Do While k > 0
                    If Not IsSpaceChar(Mid$(lt, k, 1)) Then Exit Do
                    k = k - 1
                Loop
                If k > 0 Then
                    If Mid$(lt, k, 1) = ";" Then
                        ConjunctionTailLength = Len(lt) - k + 1
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function

' "; and" is right if the next feature is the last item of its list: the last
' paragraph of the claim, or followed by a paragraph with a smaller indent.
Private Function NextIsLastOfList(ps() As Long, ByVal n As Long, ByVal k As Long) As Boolean
    If k + 1 >= n Then
        NextIsLastOfList = True
    ElseIf gIndents(ps(k + 2)) < gIndents(ps(k + 1)) - 0.5 Then
        NextIsLastOfList = True
    End If
End Function

' ---------------------------------------------------------------- spacing

Public Sub CheckSpacing(ByVal findings As Collection)
    Dim c As Long, gap As Long, p As Long, q As Long, t As String, e As Long
    Dim marks As Variant, i As Long

    For c = 2 To gClaimCount
        gap = gClaimPara(c) - gClaimLastPara(c - 1) - 1
        If gap <> 1 Then
            AddFinding findings, "spacing", gClaimPara(c), 1, Len(gClaimLabel(c)), _
                MsgSpacingBetweenClaims(gClaimLabel(c - 1), gClaimLabel(c), gap)
        End If
    Next c

    marks = Array(" .", " ,", " ;", " :")
    For p = 0 To gParaCount - 1
        t = gTexts(p)
        ' double / multiple spaces
        q = InStr(1, t, "  ")
        Do While q > 0
            e = q
            Do While e < Len(t)
                If Mid$(t, e + 1, 1) <> " " Then Exit Do
                e = e + 1
            Loop
            AddFinding findings, "spacing", p, q, e - q + 1, MsgSpacingDouble()
            q = InStr(e + 1, t, "  ")
        Loop
        ' space before punctuation (inside the claims only)
        If gClaimOfPara(p) > 0 Then
            For i = LBound(marks) To UBound(marks)
                q = InStr(1, t, marks(i))
                Do While q > 0
                    If Mid$(t, q + 2, 1) <> "." Then ' keep "..." alone
                        AddFinding findings, "spacing", p, q, 2, MsgSpacingBeforePunct(Mid$(t, q + 1, 1))
                    End If
                    q = InStr(q + 2, t, marks(i))
                Loop
            Next i
        End If
    Next p
End Sub

' ---------------------------------------------------------------- back-references

Public Sub CheckBackReferences(ByVal findings As Collection)
    Dim c As Long, r As Variant, target As Long
    For c = 1 To gClaimCount
        For Each r In gClaimRefs(c)
            target = r(R_TARGET)
            If target = 0 Then
                AddFinding findings, "backReferences", r(R_PARA), r(R_START), r(R_LEN), MsgRefUnknown(CStr(r(R_LABEL)))
            ElseIf target = c Then
                AddFinding findings, "backReferences", r(R_PARA), r(R_START), r(R_LEN), MsgRefSelf()
            ElseIf target > c Then
                AddFinding findings, "backReferences", r(R_PARA), r(R_START), r(R_LEN), MsgRefForward(CStr(r(R_LABEL)))
            End If
        Next r
    Next c
End Sub

' Independent claims (no reference in the preamble) must not refer to claims of
' the same category. References to another category are allowed, e.g.
' "... a processor configured to carry out the method of claim 1".
Public Sub CheckIndependentClaims(ByVal findings As Collection)
    Dim c As Long, r As Variant, target As Long
    For c = 1 To gClaimCount
        If Not IsDependentClaim(c) Then
            For Each r In gClaimRefs(c)
                target = r(R_TARGET)
                If target > 0 Then
                    If gClaimCategory(target) = gClaimCategory(c) Or Len(gClaimCategory(c)) = 0 _
                            Or Len(gClaimCategory(target)) = 0 Then
                        AddFinding findings, "independentClaims", r(R_PARA), r(R_KW_START), _
                            r(R_START) + r(R_LEN) - r(R_KW_START), MsgIndependentWithRef(CStr(r(R_LABEL)))
                    End If
                End If
            Next r
        End If
    Next c
End Sub

Public Sub CheckMultipleDependencies(ByVal findings As Collection)
    Dim c As Long, v As Variant, p As Long, s As Long, e As Long
    Dim isMulti() As Boolean
    If gClaimCount = 0 Then Exit Sub
    ReDim isMulti(1 To gClaimCount)
    For c = 1 To gClaimCount
        isMulti(c) = (gClaimParents(c).Count > 1)
    Next c
    For c = 1 To gClaimCount
        If isMulti(c) Then
            p = gClaimPara(c)
            s = gClaimBodyStart(c)
            e = gClaimPreambleEnd(c)
            If gOptFlagEveryMultiDep Then
                AddFinding findings, "multipleDependencies", p, s, e - s, MsgMultiDep()
            End If
            For Each v In gClaimParents(c)
                If isMulti(v) Then
                    AddFinding findings, "multipleDependencies", p, s, e - s, MsgMultiOnMulti(gClaimLabel(v))
                    Exit For
                End If
            Next v
        End If
    Next c
End Sub

' ---------------------------------------------------------------- categories

' Category named directly before a reference ("the method of claim 1",
' "des Verfahrens nach Anspruch 1"); "" if none in the 8 words before.
Public Function CategoryBeforeReference(ByVal r As Variant) As String
    Dim words As Variant, i As Long, cat As String, st As Long, n As Long
    st = r(R_KW_START) - 80
    If st < 1 Then st = 1
    words = SplitWords(Mid$(gTexts(r(R_PARA)), st, r(R_KW_START) - st))
    For i = UBound(words) To LBound(words) Step -1
        cat = CategoryOfWord(CStr(words(i)))
        If Len(cat) > 0 Then
            CategoryBeforeReference = cat
            Exit Function
        End If
        n = n + 1
        If n >= 8 Then Exit Function
    Next i
End Function

Public Sub CheckCategories(ByVal findings As Collection)
    Dim c As Long, r As Variant, target As Long, cat As String, d As Long
    For c = 1 To gClaimCount
        ' dependent claim naming another category than its parent
        For Each r In gClaimRefs(c)
            target = r(R_TARGET)
            If target > 0 And r(R_PARA) = gClaimPara(c) And r(R_KW_START) < gClaimPreambleEnd(c) Then
                cat = CategoryBeforeReference(r)
                If Len(cat) > 0 And Len(gClaimCategory(target)) > 0 And cat <> gClaimCategory(target) Then
                    AddFinding findings, "categories", r(R_PARA), r(R_KW_START), _
                        r(R_START) + r(R_LEN) - r(R_KW_START), _
                        MsgCategoryMismatch(cat, CStr(r(R_LABEL)), gClaimCategory(target))
                End If
            End If
        Next r
        ' several independent claims in the same category (Rule 43(2) EPC)
        If Not IsDependentClaim(c) And Len(gClaimCategory(c)) > 0 Then
            For d = 1 To c - 1
                If Not IsDependentClaim(d) And gClaimCategory(d) = gClaimCategory(c) Then
                    AddFinding findings, "categories", gClaimPara(c), gClaimBodyStart(c), _
                        gClaimPreambleEnd(c) - gClaimBodyStart(c), MsgCategorySecondIndependent(gClaimCategory(c), gClaimLabel(d))
                    Exit For
                End If
            Next d
        End If
    Next c
End Sub
