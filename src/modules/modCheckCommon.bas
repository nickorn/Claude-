Attribute VB_Name = "modCheckCommon"
Option Explicit

' Findings, options and dispatcher for the text checks (pure, no Word objects).
' A finding is Array(checkKey, para, start, length, message); start/length are
' positions in the visible paragraph text (length 0 = point position).

Public Const F_CHECK As Long = 0
Public Const F_PARA As Long = 1
Public Const F_START As Long = 2
Public Const F_LEN As Long = 3
Public Const F_MSG As Long = 4

' Order = order in the report and the "Einzelpr{ue}fung" menu.
Public Const CHECK_KEYS As String = "numbering|punctuation|spacing|formatting|backReferences|independentClaims|multipleDependencies|antecedentBasis|referenceSigns|blacklist|categories"

' Options (set from settings.json by the Word layer; SetDefaultCheckOptions = config defaults)
Public gOptEnforceAnd As Boolean
Public gOptFlagWherein As Boolean
Public gOptRefSignStyle As String       ' "parentheses" or "plain"
Public gOptRefSignSeparator As String   ' ", "
Public gOptNormaliseSemicolon As Boolean
Public gOptFlagEveryMultiDep As Boolean
Public gBlacklist As Collection         ' terms for the claim language

Public Sub SetDefaultCheckOptions()
    gOptEnforceAnd = True
    gOptFlagWherein = True
    gOptRefSignStyle = "parentheses"
    gOptRefSignSeparator = ", "
    gOptNormaliseSemicolon = True
    gOptFlagEveryMultiDep = False
    Set gBlacklist = New Collection
End Sub

Public Sub AddFinding(ByVal findings As Collection, ByVal checkKey As String, ByVal para As Long, _
        ByVal startPos As Long, ByVal length As Long, ByVal message As String)
    If startPos < 1 Then startPos = 1
    If length < 0 Then length = 0
    findings.Add Array(checkKey, para, startPos, length, message)
End Sub

' Runs one text check on the current claim model. Returns False for keys that
' are not text checks (e.g. "formatting", which needs Word).
Public Function RunTextCheck(ByVal checkKey As String, ByVal findings As Collection) As Boolean
    RunTextCheck = True
    Select Case checkKey
        Case "numbering": CheckNumbering findings
        Case "punctuation": CheckPunctuation findings
        Case "spacing": CheckSpacing findings
        Case "backReferences": CheckBackReferences findings
        Case "independentClaims": CheckIndependentClaims findings
        Case "multipleDependencies": CheckMultipleDependencies findings
        Case "antecedentBasis": CheckAntecedentBasis findings
        Case "referenceSigns": CheckReferenceSigns findings
        Case "blacklist": CheckBlacklist findings
        Case "categories": CheckCategories findings
        Case Else: RunTextCheck = False
    End Select
End Function

' Claim label for display in comments and the report ("-" outside the claims).
Public Function ClaimLabelOfPara(ByVal p As Long) As String
    If gClaimOfPara(p) = 0 Then
        ClaimLabelOfPara = "-"
    Else
        ClaimLabelOfPara = gClaimLabel(gClaimOfPara(p))
    End If
End Function

' Paragraph text without trailing spaces.
Public Function RTrimText(ByVal s As String) As String
    Do While Len(s) > 0
        If Not IsSpaceChar(Right$(s, 1)) Then Exit Do
        s = Left$(s, Len(s) - 1)
    Loop
    RTrimText = s
End Function

Public Function IsBlankText(ByVal s As String) As Boolean
    IsBlankText = (Len(RTrimText(s)) = 0)
End Function

' Sorts findings by (paragraph, start) ascending; returns an array.
Public Function SortFindings(ByVal findings As Collection) As Variant
    Dim arr() As Variant, i As Long, j As Long, cur As Variant
    If findings.Count = 0 Then
        SortFindings = Array()
        Exit Function
    End If
    ReDim arr(1 To findings.Count)
    For i = 1 To findings.Count
        arr(i) = findings(i)
    Next i
    For i = 2 To findings.Count
        cur = arr(i)
        j = i - 1
        Do While j >= 1
            If Not FindingIsAfter(arr(j), cur) Then Exit Do
            arr(j + 1) = arr(j)
            j = j - 1
        Loop
        arr(j + 1) = cur
    Next i
    SortFindings = arr
End Function

Private Function FindingIsAfter(ByVal a As Variant, ByVal b As Variant) As Boolean
    If a(F_PARA) <> b(F_PARA) Then
        FindingIsAfter = (a(F_PARA) > b(F_PARA))
    Else
        FindingIsAfter = (a(F_START) > b(F_START))
    End If
End Function

' Count of findings for one check key (used by tests and the summary).
Public Function CountFindings(ByVal findings As Collection, ByVal checkKey As String) As Long
    Dim f As Variant
    For Each f In findings
        If f(F_CHECK) = checkKey Then CountFindings = CountFindings + 1
    Next f
End Function
