Attribute VB_Name = "PT_ImportModules"
Option Explicit

' Developer helper - keep this module in Normal.dotm, NOT in PatentTools.dotm.
'
' Imports every module from the repository's src\ folder (modules\*.bas,
' classes\*.cls, forms\*.frm) into the open PatentTools.dotm and replaces older
' versions with the same name. Afterwards save PatentTools.dotm (Ctrl+S).
'
' One-time Word setting required:
'   File > Options > Trust Center > Trust Center Settings > Macro Settings >
'   "Trust access to the VBA project object model"

Public Sub PT_ImportModules()
    Dim target As Document, d As Document, srcRoot As String, n As Long
    For Each d In Documents
        If LCase$(d.Name) = "patenttools.dotm" Then Set target = d
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

    On Error GoTo Fail
    n = ImportFolder(target.VBProject, srcRoot & "\modules", "*.bas") _
      + ImportFolder(target.VBProject, srcRoot & "\classes", "*.cls") _
      + ImportFolder(target.VBProject, srcRoot & "\forms", "*.frm")
    MsgBox n & " Module importiert. Jetzt PatentTools.dotm speichern (Strg+S).", vbInformation, "Patent-Tools Import"
    Exit Sub
Fail:
    MsgBox "Import fehlgeschlagen: " & Err.Description & vbCr & vbCr & _
        "Ist 'Zugriff auf das VBA-Projektobjektmodell vertrauen' aktiviert?", vbCritical, "Patent-Tools Import"
End Sub

Private Function ImportFolder(ByVal proj As Object, ByVal folder As String, ByVal pattern As String) As Long
    Dim files As New Collection, f As Variant, name As String, comp As Object
    If Len(Dir$(folder, vbDirectory)) = 0 Then Exit Function
    f = Dir$(folder & "\" & pattern)
    Do While Len(f) > 0
        files.Add f
        f = Dir$()
    Loop
    For Each f In files
        name = Left$(f, InStrRev(f, ".") - 1)
        Set comp = Nothing
        On Error Resume Next
        Set comp = proj.VBComponents(name)
        On Error GoTo 0
        If Not comp Is Nothing Then
            comp.Name = name & "_old" ' free the name at once; removal may be deferred
            proj.VBComponents.Remove comp
        End If
        proj.VBComponents.Import folder & "\" & f
        ImportFolder = ImportFolder + 1
    Next f
End Function
