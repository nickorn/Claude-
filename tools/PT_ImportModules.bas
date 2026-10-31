Attribute VB_Name = "PT_ImportModules"
Option Explicit

' Developer helper - keep this module in Normal.dotm, NOT in PatentTools.dotm.
'
' Imports every module from the repository's src\ folder (modules\*.bas,
' classes\*.cls, forms\*.frm) into the open PatentTools.dotm, replaces older
' versions with the same name and removes mod* modules that are no longer in
' src (e.g. test modules). Afterwards compile and save PatentTools.dotm.
'
' One-time Word setting required:
'   File > Options > Trust Center > Trust Center Settings > Macro Settings >
'   "Trust access to the VBA project object model"

Public Sub PT_ImportModules()
    Dim target As Document, d As Document, srcRoot As String, n As Long, removed As String
    For Each d In Documents
        ' also accepts renamed downloads such as "PatentTools (1).dotm"
        If LCase$(d.Name) Like "patenttools*.dotm" Then Set target = d
    Next d
    If target Is Nothing Then
        MsgBox "Bitte zuerst PatentTools.dotm " & ChrW(252) & "ber Datei > " & ChrW(214) & "ffnen " & _
            ChrW(246) & "ffnen (nicht per Doppelklick).", vbExclamation, "Patent-Tools Import"
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Ordner 'src' des Patent-Tools-Repositorys w" & ChrW(228) & "hlen"
        If .Show <> -1 Then Exit Sub
        srcRoot = .SelectedItems(1)
    End With
    ' accept the repository root, src or src\modules as well
    If Len(Dir$(srcRoot & "\src\modules", vbDirectory)) > 0 Then srcRoot = srcRoot & "\src"
    If Len(Dir$(srcRoot & "\modules", vbDirectory)) = 0 And Len(Dir$(srcRoot & "\modClaimParse.bas")) > 0 Then
        srcRoot = Left$(srcRoot, InStrRev(srcRoot, "\") - 1)
    End If
    If Len(Dir$(srcRoot & "\modules\modClaimParse.bas")) = 0 Then
        MsgBox "Im gew" & ChrW(228) & "hlten Ordner wurde src\modules nicht gefunden:" & vbCr & srcRoot, vbExclamation, "Patent-Tools Import"
        Exit Sub
    End If

    On Error GoTo Fail
    removed = RemoveStaleModules(target.VBProject, srcRoot & "\modules")
    n = ImportFolder(target.VBProject, srcRoot & "\modules", "*.bas") _
      + ImportFolder(target.VBProject, srcRoot & "\classes", "*.cls") _
      + ImportFolder(target.VBProject, srcRoot & "\forms", "*.frm")
    MsgBox n & " Module importiert." & IIf(Len(removed) > 0, vbCr & "Entfernt (nicht mehr in src): " & removed, "") & vbCr & vbCr & _
        "Jetzt Alt+F11 > Debuggen > Kompilieren, dann PatentTools.dotm speichern (Strg+S).", vbInformation, "Patent-Tools Import"
    Exit Sub
Fail:
    MsgBox "Import fehlgeschlagen: " & Err.Description & vbCr & vbCr & _
        "Ist 'Zugriff auf das VBA-Projektobjektmodell vertrauen' aktiviert?", vbCritical, "Patent-Tools Import"
End Sub

' Standard modules named mod* that no longer exist in src\modules (renamed in a
' newer version, or test modules imported for a test run) are removed.
Private Function RemoveStaleModules(ByVal proj As Object, ByVal folder As String) As String
    Dim comp As Object, stale As New Collection, v As Variant, s As String
    For Each comp In proj.VBComponents
        If comp.Type = 1 And LCase$(comp.Name) Like "mod*" Then
            If Len(Dir$(folder & "\" & comp.Name & ".bas")) = 0 Then stale.Add comp
        End If
    Next comp
    For Each v In stale
        s = s & ", " & v.Name
        v.Name = v.Name & "_old"
        proj.VBComponents.Remove v
    Next v
    If Len(s) > 0 Then RemoveStaleModules = Mid$(s, 3)
End Function

Private Function ImportFolder(ByVal proj As Object, ByVal folder As String, ByVal pattern As String) As Long
    Dim files As New Collection, f As Variant, compName As String, comp As Object
    If Len(Dir$(folder, vbDirectory)) = 0 Then Exit Function
    f = Dir$(folder & "\" & pattern)
    Do While Len(f) > 0
        files.Add f
        f = Dir$()
    Loop
    For Each f In files
        compName = Left$(f, InStrRev(f, ".") - 1)
        Set comp = Nothing
        On Error Resume Next
        Set comp = proj.VBComponents(compName)
        On Error GoTo 0
        If Not comp Is Nothing Then
            comp.Name = compName & "_old" ' free the name at once; removal may be deferred
            proj.VBComponents.Remove comp
        End If
        proj.VBComponents.Import folder & "\" & f
        ImportFolder = ImportFolder + 1
    Next f
End Function
