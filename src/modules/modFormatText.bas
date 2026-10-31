Attribute VB_Name = "modFormatText"
Option Explicit

' MVP-3 text fixes, planned on the claim model (pure, no Word objects).
' Each planner adds replace operations (modClaimParse OP_* records, kind
' OPK_REFERENCE = replace; length 0 = insertion) for modFormat to apply.

' ---------------------------------------------------------------- spaces

' Multiple spaces -> one space (whole document); spaces before . , ; : and at
' the end of a paragraph are removed (inside the claims).
Public Sub PlanSpaceFixes(ByVal ops As Collection)
    Dim p As Long, t As String, q As Long, e As Long, nextChar As String, inClaims As Boolean
    For p = 0 To gParaCount - 1
        t = gTexts(p)
        inClaims = (gClaimOfPara(p) > 0)
        q = InStr(1, t, " ")
        Do While q > 0
            e = q
            Do While e < Len(t)
                If Mid$(t, e + 1, 1) <> " " Then Exit Do
                e = e + 1
            Loop
            nextChar = Mid$(t, e + 1, 1)
            If inClaims And (nextChar = "" Or nextChar = "." Or nextChar = "," Or nextChar = ";" Or nextChar = ":") _
                    And q > 1 And Mid$(t, e + 1, 3) <> "..." Then
                ops.Add Array(OPK_REFERENCE, p, q, e - q + 1, "")
            ElseIf e > q Then
                ops.Add Array(OPK_REFERENCE, p, q, e - q + 1, " ")
            End If
            q = InStr(e + 1, t, " ")
        Loop
    Next p
End Sub

' ---------------------------------------------------------------- punctuation

' Features end with ";" (":" for the preamble of a multi-paragraph claim), the
' last feature with ".", "; and"/"; und" before the last feature (setting).
' Claims with text after an inner full stop are only fixed up to that stop;
' misplaced "; and" is left alone (see Vollpruefung).
Public Sub PlanPunctuationFixes(ByVal ops As Collection)
    Dim c As Long, p As Long, n As Long, k As Long, listEnd As Long, t As String, lastCh As String
    Dim ps() As Long, conj As String
    If gLanguage = "de" Then conj = " und" Else conj = " and"
    For c = 1 To gClaimCount
        ReDim ps(1 To gClaimLastPara(c) - gClaimPara(c) + 1)
        n = 0
        For p = gClaimPara(c) To gClaimLastPara(c)
            If Not IsBlankText(gTexts(p)) Then
                n = n + 1
                ps(n) = p
            End If
        Next p
        listEnd = n
        For k = 1 To n - 1
            If Right$(RTrimText(gTexts(ps(k))), 1) = "." Then
                listEnd = k
                Exit For
            End If
        Next k

        For k = 1 To listEnd - 1
            t = RTrimText(gTexts(ps(k)))
            lastCh = Right$(t, 1)
            If lastCh = ":" Or HasConjunctionTail(t) Then
                ' fine as it is
            ElseIf lastCh = ";" Then
                If NeedsConjunction(k, listEnd) Then ops.Add Array(OPK_REFERENCE, ps(k), Len(t) + 1, 0, conj)
            ElseIf k = 1 Then
                If lastCh = "," Then
                    ops.Add Array(OPK_REFERENCE, ps(k), Len(t), 1, ":")
                Else
                    ops.Add Array(OPK_REFERENCE, ps(k), Len(t) + 1, 0, ":")
                End If
            ElseIf lastCh = "," Then
                ops.Add Array(OPK_REFERENCE, ps(k), Len(t), 1, ";" & IIf(NeedsConjunction(k, listEnd), conj, ""))
            ElseIf Len(t) > 0 Then
                ops.Add Array(OPK_REFERENCE, ps(k), Len(t) + 1, 0, ";" & IIf(NeedsConjunction(k, listEnd), conj, ""))
            End If
        Next k

        ' end of the claim (only when nothing follows an inner full stop)
        If listEnd = n Then
            t = RTrimText(gTexts(ps(n)))
            lastCh = Right$(t, 1)
            If Right$(t, 2) = ".." And Right$(t, 3) <> "..." Then
                ops.Add Array(OPK_REFERENCE, ps(n), Len(t) - 1, 2, ".")
            ElseIf lastCh = ";" Or lastCh = "," Then
                ops.Add Array(OPK_REFERENCE, ps(n), Len(t), 1, ".")
            ElseIf lastCh <> "." And lastCh <> ":" And Len(t) > 0 Then
                ops.Add Array(OPK_REFERENCE, ps(n), Len(t) + 1, 0, ".")
            End If
        End If
    Next c
End Sub

Private Function NeedsConjunction(ByVal k As Long, ByVal listEnd As Long) As Boolean
    NeedsConjunction = gOptEnforceAnd And listEnd >= 3 And k = listEnd - 1
End Function

Private Function HasConjunctionTail(ByVal t As String) As Boolean
    Dim lt As String, words As Variant, i As Long, k As Long
    lt = LCase$(t)
    words = Array(" and/or", " und/oder", " oder", " and", " und", " or")
    For i = LBound(words) To UBound(words)
        If EndsWith(lt, words(i)) Then
            k = Len(lt) - Len(words(i))
            Do While k > 0
                If Not IsSpaceChar(Mid$(lt, k, 1)) Then Exit Do
                k = k - 1
            Loop
            If k > 0 Then
                If Mid$(lt, k, 1) = ";" Then
                    HasConjunctionTail = True
                    Exit Function
                End If
            End If
        End If
    Next i
End Function

' ---------------------------------------------------------------- reference signs

' "(150; 450)" / "( 150 ,450 )" -> "(150, 450)" (separator from settings) and
' a space before the opening parenthesis ("controller(150)" -> "controller (150)").
Public Sub PlanSignFormatFixes(ByVal ops As Collection)
    Dim p As Long, t As String, q As Long, e As Long, inner As String, want As String
    Dim parts As Variant, i As Long, before As String
    If gClaimCount = 0 Then Exit Sub
    For p = gClaimPara(1) To gParaCount - 1
        t = gTexts(p)
        q = InStr(ParaBodyStart(p), t, "(")
        Do While q > 0
            e = InStr(q + 1, t, ")")
            If e = 0 Then Exit Do
            inner = Mid$(t, q + 1, e - q - 1)
            If IsSignList(inner) Then
                parts = Split(Replace(inner, ";", ","), ",")
                For i = LBound(parts) To UBound(parts)
                    parts(i) = Trim$(parts(i))
                Next i
                want = "(" & Join(parts, gOptRefSignSeparator) & ")"
                before = ""
                If q > 1 Then
                    If IsLetterChar(Mid$(t, q - 1, 1)) Or IsDigitChar(Mid$(t, q - 1, 1)) Then before = " "
                End If
                If before & want <> Mid$(t, q, e - q + 1) Then
                    ops.Add Array(OPK_REFERENCE, p, q, e - q + 1, before & want)
                End If
            End If
            q = InStr(e + 1, t, "(")
        Loop
    Next p
End Sub
