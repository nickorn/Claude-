Attribute VB_Name = "modDraftPlan"
Option Explicit

' Roadmap 6 drafting helpers, pure part (claim model in, text out):
' claim tree, dependent-claim skeleton, merge of a dependent claim into its parent.

' Tree line: Array(depth, label, excerpt, note)
Public Const TL_DEPTH As Long = 0
Public Const TL_LABEL As Long = 1
Public Const TL_TEXT As Long = 2
Public Const TL_NOTE As Long = 3

' ---------------------------------------------------------------- claim tree

' Each claim appears once, below its first parent; further parents are noted.
Public Function ClaimTreeLines() As Collection
    Dim lines As New Collection, shown() As Boolean, c As Long
    If gClaimCount = 0 Then
        Set ClaimTreeLines = lines
        Exit Function
    End If
    ReDim shown(1 To gClaimCount)
    For c = 1 To gClaimCount
        If FirstParent(c) = 0 Then AddTreeNode lines, shown, c, 0
    Next c
    ' claims only reachable through forward references or cycles
    For c = 1 To gClaimCount
        If Not shown(c) Then AddTreeNode lines, shown, c, 0
    Next c
    Set ClaimTreeLines = lines
End Function

Private Sub AddTreeNode(ByVal lines As Collection, ByRef shown() As Boolean, ByVal c As Long, ByVal depth As Long)
    Dim d As Long
    If shown(c) Then Exit Sub
    shown(c) = True
    lines.Add Array(depth, gClaimLabel(c), ClaimExcerpt(c, 70), TreeNote(c))
    For d = 1 To gClaimCount
        If FirstParent(d) = c And d <> c Then AddTreeNode lines, shown, d, depth + 1
    Next d
End Sub

' Smallest existing parent before the claim itself; 0 for independent claims.
Public Function FirstParent(ByVal c As Long) As Long
    Dim v As Variant
    For Each v In gClaimParents(c)
        If v <> c Then
            If FirstParent = 0 Or v < FirstParent Then FirstParent = v
        End If
    Next v
End Function

Private Function TreeNote(ByVal c As Long) As String
    Dim v As Variant, others As String, fp As Long
    fp = FirstParent(c)
    If fp = 0 Then
        TreeNote = CategoryName(gClaimCategory(c))
        Exit Function
    End If
    For Each v In gClaimParents(c)
        If v <> fp Then others = others & ", " & gClaimLabel(v)
    Next v
    If Len(others) > 0 Then TreeNote = Txt("auch abh{ae}ngig von ") & Mid$(others, 3)
    If fp > c Then TreeNote = Trim$(TreeNote & Txt(" (Vorw{ae}rtsbezug!)"))
End Function

Public Function ClaimExcerpt(ByVal c As Long, ByVal maxLen As Long) As String
    Dim t As String
    t = Mid$(gTexts(gClaimPara(c)), gClaimBodyStart(c))
    If Len(t) > maxLen Then t = Left$(t, maxLen) & ChrW(8230)
    ClaimExcerpt = t
End Function

' ---------------------------------------------------------------- dependent claim skeleton

' "The method according to claim 3, wherein " / "Verfahren nach Anspruch 3, wobei "
Public Function DependentSkeleton(ByVal parent As Long) As String
    Dim noun As String
    noun = CategoryNoun(parent)
    If gLanguage = "de" Then
        If Len(noun) = 0 Then noun = "Gegenstand"
        DependentSkeleton = UCase$(Left$(noun, 1)) & Mid$(noun, 2) & " nach Anspruch " & gClaimLabel(parent) & ", wobei "
    Else
        If Len(noun) = 0 Then noun = "subject-matter"
        DependentSkeleton = "The " & LCase$(noun) & " according to claim " & gClaimLabel(parent) & ", wherein "
    End If
End Function

' The word of the preamble that gives the claim its category ("method",
' "Vorrichtung", "storage medium" -> "medium"); for dependent claims the
' category word of the root claim.
Public Function CategoryNoun(ByVal c As Long) As String
    Dim words As Variant, i As Long, j As Long, guard As Long, cat As String, w As String
    Do While FirstParent(c) > 0 And FirstParent(c) < c And guard < 50
        c = FirstParent(c)
        guard = guard + 1
    Loop
    words = SplitWords(Mid$(gTexts(gClaimPara(c)), gClaimBodyStart(c), gClaimPreambleEnd(c) - gClaimBodyStart(c)))
    For i = LBound(words) To UBound(words)
        cat = CategoryOfWord(CStr(words(i)))
        If Len(cat) > 0 Then
            ' the head noun may follow ("computer-readable storage medium" -> "medium")
            CategoryNoun = words(i)
            For j = i + 1 To i + 3
                If j > UBound(words) Then Exit For
                w = CategoryOfWord(CStr(words(j)))
                If w = cat Then CategoryNoun = words(j)
                If Len(w) > 0 And w <> cat Then Exit For
            Next j
            Exit Function
        End If
    Next i
End Function

' First free supplementary label after a claim: 3 -> 3a (or 3b, ... if taken).
Public Function NextSupplementaryLabel(ByVal label As String) As String
    Dim i As Long, cand As String
    For i = 0 To 25
        cand = label & Chr$(97 + i)
        If ClaimNumberOfLabel(cand) = 0 Then
            NextSupplementaryLabel = cand
            Exit Function
        End If
    Next i
    NextSupplementaryLabel = label & "z1"
End Function

' ---------------------------------------------------------------- merge

' Plan to merge dependent claim d into its parent. Returns
' Array(ok, message, parent, tailText, removeConjPara, removeConjLen):
'   tailText replaces the final full stop of the parent's last paragraph
'   (", wherein ..." for wherein-clauses, ";" & vbCr & features for
'   "further comprising"); removeConjPara/Len: "; and" to drop from the
'   parent's old penultimate feature (-1 = none).
Public Function PlanMerge(ByVal d As Long) As Variant
    Dim parent As Long, rest As String, lt As String, extra As String, p As Long, terms As String
    Dim lastText As String, featureMode As Boolean, k As Long, words As Variant, i As Long
    Dim pen As Long, penText As String, conjPos As Long, conj As String, parents As Collection

    If d < 1 Or d > gClaimCount Then
        PlanMerge = Array(False, MsgMergeNotDependent(), 0, "", -1, 0)
        Exit Function
    End If
    If gClaimParents(d).Count <> 1 Then
        PlanMerge = Array(False, MsgMergeNotDependent(), 0, "", -1, 0)
        Exit Function
    End If
    Set parents = gClaimParents(d)
    parent = parents(1)
    If parent >= d Then
        PlanMerge = Array(False, MsgMergeForward(), 0, "", -1, 0)
        Exit Function
    End If
    lastText = RTrimText(gTexts(gClaimLastPara(parent)))
    If Right$(lastText, 1) <> "." Then
        PlanMerge = Array(False, MsgMergeParentEnd(gClaimLabel(parent)), 0, "", -1, 0)
        Exit Function
    End If

    ' text of d after its preamble ("The method of claim 1")
    rest = Mid$(gTexts(gClaimPara(d)), gClaimPreambleEnd(d))
    Do While Len(rest) > 0
        If Left$(rest, 1) <> "," And Not IsSpaceChar(Left$(rest, 1)) Then Exit Do
        rest = Mid$(rest, 2)
    Loop
    For p = gClaimPara(d) + 1 To gClaimLastPara(d)
        If Not IsBlankText(gTexts(p)) Then extra = extra & vbCr & RTrimText(gTexts(p))
    Next p
    terms = TermsOfClaimChain(parent)
    rest = AdjustArticles(RTrimText(rest), terms)
    extra = AdjustArticles(extra, terms)
    If Len(rest) = 0 And Len(extra) = 0 Then
        PlanMerge = Array(False, MsgMergeEmpty(), 0, "", -1, 0)
        Exit Function
    End If

    ' "further comprising ..." adds features; anything else is a wherein-clause
    lt = LCase$(rest)
    words = Array("further comprising", "further including", "comprising", "including", _
        "ferner umfassend", "ferner aufweisend", "umfassend", "aufweisend", "ferner mit", "mit")
    For i = LBound(words) To UBound(words)
        If Left$(lt, Len(words(i))) = words(i) Then
            k = Len(words(i)) + 1
            If Mid$(rest, k, 1) = ":" Then k = k + 1
            rest = Trim$(Mid$(rest, k))
            featureMode = True
            Exit For
        End If
    Next i

    If gLanguage = "de" Then conj = " und" Else conj = " and"
    If featureMode Then
        ' parent "...last feature." -> "...last feature; and" + new feature(s)
        pen = -1
        PlanMerge = Array(True, "", parent, ";" & IIf(gOptEnforceAnd, conj, "") & _
            IIf(Len(rest) > 0, vbCr & rest, "") & extra, -1, 0)
        If gOptEnforceAnd Then
            pen = PreviousFeaturePara(parent)
            If pen >= 0 Then
                penText = RTrimText(gTexts(pen))
                conjPos = InStrRev(LCase$(penText), ";")
                If conjPos > 0 Then
                    If LCase$(Trim$(Mid$(penText, conjPos + 1))) = Trim$(conj) Then
                        PlanMerge = Array(True, "", parent, ";" & conj & IIf(Len(rest) > 0, vbCr & rest, "") & extra, _
                            pen, Len(penText) - conjPos)
                    End If
                End If
            End If
        End If
    Else
        PlanMerge = Array(True, "", parent, ", " & rest & extra, -1, 0)
    End If
End Function

' Non-empty paragraph before the parent's last paragraph (inside the claim), -1 if none.
Private Function PreviousFeaturePara(ByVal c As Long) As Long
    Dim p As Long
    PreviousFeaturePara = -1
    For p = gClaimLastPara(c) - 1 To gClaimPara(c) Step -1
        If Not IsBlankText(gTexts(p)) Then
            PreviousFeaturePara = p
            Exit Function
        End If
    Next p
End Function
