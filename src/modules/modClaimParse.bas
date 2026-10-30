Attribute VB_Name = "modClaimParse"
Option Explicit

' Pure text logic for claim labels and back-references.
' No Word objects in here, so everything can be unit-tested on plain strings
' (tests/vba/modTestRenumber.bas, also runnable in LibreOffice Basic).
'
' Positions are 1-based character positions within one paragraph text.
' A claim paragraph starts with  <label> "." TAB  where label matches
' ^\d+[0-9A-Za-z.]*$ and is treated as an opaque token (1a1 <> 1a.1).

' Planned operation: Array(kind, paraIndex, startPos, length, text)
Public Const OP_KIND As Long = 0
Public Const OP_PARA As Long = 1
Public Const OP_START As Long = 2
Public Const OP_LEN As Long = 3
Public Const OP_TEXT As Long = 4

Public Const OPK_RELABEL As Long = 0   ' replace a claim's own label by text
Public Const OPK_REFERENCE As Long = 1 ' replace a back-reference label by text
Public Const OPK_FLAG As Long = 2      ' attach a comment (text) to the span

' Reference token: Array(startPos, length, label, isRangeEnd, keywordStart, keywordLength)
Public Const TK_START As Long = 0
Public Const TK_LEN As Long = 1
Public Const TK_LABEL As Long = 2
Public Const TK_RANGE_END As Long = 3  ' True: closes a range opened by the previous token
Public Const TK_KW_START As Long = 4
Public Const TK_KW_LEN As Long = 5

Private Const KW_EN As String = "claim"
Private Const KW_DE_SINGULAR As String = "anspruch"

' ---------------------------------------------------------------- characters

Private Function IsDigitChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    Dim a As Long
    a = AscW(c)
    IsDigitChar = (a >= 48 And a <= 57)
End Function

Private Function IsLabelChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    Dim a As Long
    a = AscW(c)
    IsLabelChar = (a >= 48 And a <= 57) Or (a >= 65 And a <= 90) Or (a >= 97 And a <= 122) Or a = 46
End Function

Private Function IsLetterChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    IsLetterChar = (LCase$(c) <> UCase$(c)) Or AscW(c) = 223
End Function

Private Function IsSpaceChar(ByVal c As String) As Boolean
    If Len(c) <> 1 Then Exit Function
    Select Case AscW(c)
        Case 32, 9, 160, 8194, 8195, 8201, 8239
            IsSpaceChar = True
    End Select
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsDigitChar(Mid$(s, i, 1)) Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function SkipSpaces(ByVal s As String, ByVal pos As Long) As Long
    Do While pos <= Len(s)
        If Not IsSpaceChar(Mid$(s, pos, 1)) Then Exit Do
        pos = pos + 1
    Loop
    SkipSpaces = pos
End Function

' ---------------------------------------------------------------- claim labels

' True if text is a claim paragraph; returns its label and the position where
' the claim body starts (after the TAB).
Public Function ParseClaimLabel(ByVal paraText As String, ByRef label As String, ByRef bodyStart As Long) As Boolean
    Dim i As Long
    If Len(paraText) < 3 Then Exit Function
    If Not IsDigitChar(Left$(paraText, 1)) Then Exit Function
    i = 2
    Do While i <= Len(paraText)
        If Not IsLabelChar(Mid$(paraText, i, 1)) Then Exit Do
        i = i + 1
    Loop
    ' text(1 .. i-1) is the label run; it must end with "." and be followed by TAB
    If i > Len(paraText) Then Exit Function
    If Mid$(paraText, i, 1) <> vbTab Then Exit Function
    If Mid$(paraText, i - 1, 1) <> "." Then Exit Function
    label = Left$(paraText, i - 2)
    bodyStart = i + 1
    ParseClaimLabel = True
End Function

' Collection keys are case-insensitive; labels are not (1a <> 1A).
Private Function LabelKey(ByVal label As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(label)
        c = Mid$(label, i, 1)
        If c >= "A" And c <= "Z" Then
            r = r & "^" & c
        Else
            r = r & c
        End If
    Next i
    LabelKey = "k" & r
End Function

' Returns the 1-based claim position for a label, 0 if unknown.
Private Function LookupClaim(ByVal index As Collection, ByVal label As String) As Long
    On Error Resume Next
    LookupClaim = index.Item(LabelKey(label))
    If Err.Number <> 0 Then LookupClaim = 0
    Err.Clear
End Function

' ---------------------------------------------------------------- back-references

Private Function KwDePlural() As String
    KwDePlural = "anspr" & ChrW(252) & "ch"
End Function

' If a keyword starts at q in lt (lower-case text), returns the position after
' keyword + inflection; 0 if this occurrence is not a claim keyword.
Private Function KeywordEnd(ByVal lt As String, ByVal q As Long, ByVal kw As String) As Long
    Dim e As Long
    e = q + Len(kw)
    If kw = KW_EN Then
        ' English: whole word only ("disclaim", "claimed" do not count)
        If q > 1 Then
            If IsLetterChar(Mid$(lt, q - 1, 1)) Then Exit Function
        End If
        If Mid$(lt, e, 1) = "s" Then e = e + 1
    ElseIf kw = KW_DE_SINGULAR Then
        ' German compounds allowed before ("Patentanspruch"); Anspruch/Anspruchs/Anspruches
        If Mid$(lt, e, 2) = "es" Then
            e = e + 2
        ElseIf Mid$(lt, e, 1) = "s" Then
            e = e + 1
        End If
    Else
        ' Ansprueche / Anspruechen
        If Mid$(lt, e, 1) <> "e" Then Exit Function
        e = e + 1
        If Mid$(lt, e, 1) = "n" Then e = e + 1
    End If
    If IsLetterChar(Mid$(lt, e, 1)) Then Exit Function
    KeywordEnd = e
End Function

Private Function FindNextKeyword(ByVal lt As String, ByVal startPos As Long, _
        ByRef kwPos As Long, ByRef kwEnd As Long) As Boolean
    Dim kws(2) As String, i As Long, q As Long, e As Long, s As Long
    kws(0) = KW_EN
    kws(1) = KW_DE_SINGULAR
    kws(2) = KwDePlural()
    kwPos = 0
    For i = 0 To 2
        s = startPos
        Do
            q = InStr(s, lt, kws(i))
            If q = 0 Then Exit Do
            If kwPos > 0 And q >= kwPos Then Exit Do
            e = KeywordEnd(lt, q, kws(i))
            If e > 0 Then
                kwPos = q
                kwEnd = e
                Exit Do
            End If
            s = q + 1
        Loop
    Next i
    FindNextKeyword = (kwPos > 0)
End Function

' Parses a list separator after a label. Returns the position of the next
' token (0 = no separator). isRange is True for "to"/"through"/"bis"/dashes.
Private Function ParseSeparator(ByVal lt As String, ByVal pos As Long, ByRef isRange As Boolean) As Long
    Dim k As Long, i As Long, w As String, c As String, hadComma As Boolean
    Dim words As Variant
    isRange = False
    k = SkipSpaces(lt, pos)
    If Mid$(lt, k, 1) = "," Then
        hadComma = True
        k = SkipSpaces(lt, k + 1)
    End If
    c = Mid$(lt, k, 1)
    If c = "-" Or c = ChrW(8211) Or c = ChrW(8212) Then
        isRange = True
        ParseSeparator = SkipSpaces(lt, k + 1)
        Exit Function
    End If
    ' longest first, so "and/or" wins over "and"
    words = Array("and/or", "und/oder", "through", "sowie", "oder", "und", "and", "bis", "to", "or")
    For i = LBound(words) To UBound(words)
        w = words(i)
        If Mid$(lt, k, Len(w)) = w Then
            If Not IsLetterChar(Mid$(lt, k + Len(w), 1)) Then
                isRange = (w = "to" Or w = "through" Or w = "bis")
                ParseSeparator = SkipSpaces(lt, k + Len(w))
                Exit Function
            End If
        End If
    Next i
    If hadComma Then ParseSeparator = k
End Function

' Parses "1a, 2 or 3 to 5" starting at j (a digit). Adds tokens to refs and
' returns the position after the last label.
Private Function ParseReferenceList(ByVal paraText As String, ByVal lt As String, ByVal j As Long, _
        ByVal kwStart As Long, ByVal kwLen As Long, ByVal refs As Collection) As Long
    Dim st As Long, en As Long, k As Long
    Dim isRange As Boolean, pendingRange As Boolean
    Do
        st = j
        en = j
        Do While en <= Len(paraText)
            If Not IsLabelChar(Mid$(paraText, en, 1)) Then Exit Do
            en = en + 1
        Loop
        ' a trailing full stop ends the sentence, it is not part of the label
        Do While en - 1 > st
            If Mid$(paraText, en - 1, 1) <> "." Then Exit Do
            en = en - 1
        Loop
        refs.Add Array(st, en - st, Mid$(paraText, st, en - st), pendingRange, kwStart, kwLen)
        k = ParseSeparator(lt, en, isRange)
        If k = 0 Then Exit Do
        If Not IsDigitChar(Mid$(lt, k, 1)) Then Exit Do
        pendingRange = isRange
        j = k
    Loop
    ParseReferenceList = en
End Function

' Finds all back-references (EN "claim(s) ...", DE "Anspruch/Ansprueche ...")
' in text from fromPos on. Returns a Collection of reference tokens (TK_*).
Public Function FindReferences(ByVal paraText As String, ByVal fromPos As Long) As Collection
    Dim refs As New Collection
    Dim lt As String, p As Long, kwPos As Long, kwEnd As Long, j As Long
    lt = LCase$(paraText)
    p = fromPos
    Do While p <= Len(lt)
        If Not FindNextKeyword(lt, p, kwPos, kwEnd) Then Exit Do
        j = SkipSpaces(lt, kwEnd)
        If j > kwEnd And IsDigitChar(Mid$(lt, j, 1)) Then
            p = ParseReferenceList(paraText, lt, j, kwPos, kwEnd - kwPos, refs)
        Else
            p = kwEnd
        End If
    Loop
    Set FindReferences = refs
End Function

' ---------------------------------------------------------------- planning

' A range "X to Y" is rewritten by document position. It is safe when the
' new range covers exactly the claims the old integer range covered.
Private Function RangeIsSafe(ByVal oldFrom As String, ByVal oldTo As String, _
        ByVal newFrom As Long, ByVal newTo As Long, ByVal index As Collection) As Boolean
    Dim x As Long, y As Long, m As Long, k As Long
    If newFrom > newTo Then Exit Function
    If IsAllDigits(oldFrom) And IsAllDigits(oldTo) And Len(oldFrom) <= 6 And Len(oldTo) <= 6 Then
        x = CLng(oldFrom)
        y = CLng(oldTo)
        If y - x <> newTo - newFrom Then Exit Function
        For m = x + 1 To y - 1
            k = LookupClaim(index, CStr(m))
            If k > 0 Then
                If k < newFrom Or k > newTo Then Exit Function
            End If
        Next m
    End If
    RangeIsSafe = True
End Function

Private Sub PlanReference(ByVal ops As Collection, ByVal para As Long, ByVal t As Variant, _
        ByVal prev As Variant, ByVal hasPrev As Boolean, ByVal index As Collection)
    Dim a As Long, b As Long
    b = LookupClaim(index, CStr(t(TK_LABEL)))
    If b = 0 Then
        ops.Add Array(OPK_FLAG, para, t(TK_START), t(TK_LEN), CmtUnknownReference(CStr(t(TK_LABEL))))
        Exit Sub
    End If
    If CStr(b) <> CStr(t(TK_LABEL)) Then
        ops.Add Array(OPK_REFERENCE, para, t(TK_START), t(TK_LEN), CStr(b))
    End If
    If Not hasPrev Then Exit Sub
    If Not t(TK_RANGE_END) Then Exit Sub
    a = LookupClaim(index, CStr(prev(TK_LABEL)))
    If a = 0 Then Exit Sub ' already flagged as unknown
    If Not RangeIsSafe(CStr(prev(TK_LABEL)), CStr(t(TK_LABEL)), a, b, index) Then
        ops.Add Array(OPK_FLAG, para, t(TK_KW_START), t(TK_KW_LEN), _
            CmtRangeCheck(CStr(prev(TK_LABEL)), CStr(t(TK_LABEL)), a, b))
    End If
End Sub

' Plans the renumbering for the paragraph texts (0- or 1-based Variant array
' of visible paragraph texts without the paragraph mark).
' New number = position of the claim in the document. Fills ops with OP_*
' records in document order. Returns False (and errMsg) if nothing can be done.
Public Function PlanRenumber(ByVal texts As Variant, ByVal ops As Collection, _
        ByRef errMsg As String, ByRef claimCount As Long) As Boolean
    Dim i As Long, n As Long, lo As Long, hi As Long, claimNo As Long
    Dim lbl As String, bodyStart As Long, dupes As String, firstPara As Long, found As Boolean
    Dim isClaim() As Boolean, bodyOf() As Long
    Dim index As New Collection
    Dim refs As Collection, t As Variant, prev As Variant, hasPrev As Boolean

    lo = LBound(texts)
    hi = UBound(texts)
    claimCount = 0
    If hi < lo Then
        errMsg = MsgNoClaims()
        Exit Function
    End If
    ReDim isClaim(lo To hi)
    ReDim bodyOf(lo To hi)

    ' 1. collect labels in document order; duplicates abort
    For i = lo To hi
        If ParseClaimLabel(CStr(texts(i)), lbl, bodyStart) Then
            If LookupClaim(index, lbl) > 0 Then
                If InStr(1, Chr$(1) & dupes & Chr$(1), Chr$(1) & lbl & Chr$(1)) = 0 Then
                    If Len(dupes) > 0 Then dupes = dupes & Chr$(1)
                    dupes = dupes & lbl
                End If
            Else
                n = n + 1
                index.Add n, LabelKey(lbl)
            End If
            isClaim(i) = True
            bodyOf(i) = bodyStart
            If Not found Then
                firstPara = i
                found = True
            End If
        End If
    Next i
    If Len(dupes) > 0 Then
        errMsg = MsgDuplicateLabels(Replace(dupes, Chr$(1), ", "))
        Exit Function
    End If
    If n = 0 Then
        errMsg = MsgNoClaims()
        Exit Function
    End If
    claimCount = n

    ' 2. relabel claims and rewrite references (text before claim 1 is ignored)
    For i = firstPara To hi
        If isClaim(i) Then
            claimNo = claimNo + 1
            ParseClaimLabel CStr(texts(i)), lbl, bodyStart
            If lbl <> CStr(claimNo) Then
                ops.Add Array(OPK_RELABEL, i, 1, Len(lbl), CStr(claimNo))
            End If
        Else
            bodyStart = 1
        End If
        Set refs = FindReferences(CStr(texts(i)), bodyStart)
        hasPrev = False
        For Each t In refs
            PlanReference ops, i, t, prev, hasPrev, index
            prev = t
            hasPrev = True
        Next t
    Next i
    PlanRenumber = True
End Function

' ---------------------------------------------------------------- applying

' Returns the ops as an array sorted by (paragraph, start position), so callers
' can apply them from the last to the first and keep earlier positions valid.
Public Function SortOps(ByVal ops As Collection) As Variant
    Dim arr() As Variant, i As Long, j As Long, cur As Variant
    If ops.Count = 0 Then
        SortOps = Array()
        Exit Function
    End If
    ReDim arr(1 To ops.Count)
    For i = 1 To ops.Count
        arr(i) = ops(i)
    Next i
    For i = 2 To ops.Count
        cur = arr(i)
        j = i - 1
        Do While j >= 1
            If Not OpIsAfter(arr(j), cur) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = cur
    Next i
    SortOps = arr
End Function

Private Function OpIsAfter(ByVal a As Variant, ByVal b As Variant) As Boolean
    If a(OP_PARA) <> b(OP_PARA) Then
        OpIsAfter = (a(OP_PARA) > b(OP_PARA))
    Else
        OpIsAfter = (a(OP_START) > b(OP_START))
    End If
End Function

' Applies ops to plain texts (used by the unit tests). Flags are returned as
' "para:start:text" lines in flagLog.
Public Function ApplyOpsToTexts(ByVal texts As Variant, ByVal ops As Collection, ByRef flagLog As String) As Variant
    Dim sorted As Variant, i As Long, op As Variant, s As String
    sorted = SortOps(ops)
    flagLog = ""
    For i = UBound(sorted) To LBound(sorted) Step -1
        op = sorted(i)
        If op(OP_KIND) = OPK_FLAG Then
            flagLog = op(OP_PARA) & ":" & op(OP_START) & ":" & op(OP_TEXT) & vbLf & flagLog
        Else
            s = texts(op(OP_PARA))
            texts(op(OP_PARA)) = Left$(s, op(OP_START) - 1) & op(OP_TEXT) & Mid$(s, op(OP_START) + op(OP_LEN))
        End If
    Next i
    ApplyOpsToTexts = texts
End Function
