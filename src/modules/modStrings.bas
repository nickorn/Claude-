Attribute VB_Name = "modStrings"
Option Explicit

' Central place for all German UI texts (message boxes, comments, undo titles).
' Ribbon labels live in src/customUI/customUI14.xml.
'
' Source files must stay pure ASCII: the VBA editor imports .bas files in the
' ANSI code page, so umlauts would break. Write them as tokens instead:
'   {ae} {oe} {ue} {Ae} {Oe} {Ue} {ss}   umlauts / sharp s
'   {-}                                  en dash
'   {q<} {q>}                            German quotes (low / high)

Public Const APP_VERSION As String = "0.2.0"

Public Function AppTitle() As String
    AppTitle = "Patent-Tools"
End Function

' Expands the umlaut/typography tokens described above.
Public Function Txt(ByVal s As String) As String
    s = Replace(s, "{ae}", ChrW(228))
    s = Replace(s, "{oe}", ChrW(246))
    s = Replace(s, "{ue}", ChrW(252))
    s = Replace(s, "{Ae}", ChrW(196))
    s = Replace(s, "{Oe}", ChrW(214))
    s = Replace(s, "{Ue}", ChrW(220))
    s = Replace(s, "{ss}", ChrW(223))
    s = Replace(s, "{-}", ChrW(8211))
    s = Replace(s, "{q<}", ChrW(8222))
    s = Replace(s, "{q>}", ChrW(8220))
    Txt = s
End Function

Private Function Quoted(ByVal s As String) As String
    Quoted = Txt("{q<}") & s & Txt("{q>}")
End Function

' ---------------------------------------------------------------- general

Public Function MsgNoDocument() As String
    MsgNoDocument = Txt("Es ist kein Dokument ge{oe}ffnet.")
End Function

Public Function MsgUnexpectedError(ByVal description As String) As String
    MsgUnexpectedError = Txt("Unerwarteter Fehler: ") & description & vbCr & vbCr & _
        Txt("Bereits vorgenommene {Ae}nderungen k{oe}nnen mit Strg+Z r{ue}ckg{ae}ngig gemacht werden.")
End Function

Public Function MsgInfo(ByVal folder As String) As String
    MsgInfo = AppTitle() & " Version " & APP_VERSION & vbCr & vbCr & _
        Txt("Einstellungen: ") & folder & vbCr & _
        Txt("Offline-Werkzeug {-} keine Netzwerkzugriffe.")
End Function

Public Function MsgSettingsError(ByVal path As String, ByVal description As String) As String
    MsgSettingsError = Txt("Die Einstellungsdatei konnte nicht gelesen werden:") & vbCr & path & vbCr & vbCr & _
        description & vbCr & vbCr & Txt("Es werden die Standardwerte verwendet.")
End Function

Public Function MsgConfigFolderError(ByVal path As String, ByVal description As String) As String
    MsgConfigFolderError = Txt("Der Einstellungsordner konnte nicht angelegt werden:") & vbCr & path & vbCr & vbCr & description
End Function

' ---------------------------------------------------------------- renumbering

Public Function UndoTitleRenumber() As String
    UndoTitleRenumber = AppTitle() & ": Neu nummerieren"
End Function

Public Function MsgNoClaims() As String
    MsgNoClaims = Txt("Keine Anspr{ue}che gefunden.") & vbCr & vbCr & _
        Txt("Erwartet wird je Anspruch ein Absatz, der mit Nummer, Punkt und Tabulator beginnt (z. B. ") & _
        Quoted("1.<Tab>") & ")."
End Function

Public Function MsgDuplicateLabels(ByVal labelList As String) As String
    MsgDuplicateLabels = Txt("Doppelte Anspruchsnummern gefunden: ") & labelList & vbCr & vbCr & _
        Txt("Die Neunummerierung wurde abgebrochen, das Dokument ist unver{ae}ndert. ") & _
        Txt("Bitte die Nummern zuerst eindeutig machen.")
End Function

Public Function MsgAlreadyConsecutive(ByVal claimCount As Long) As String
    MsgAlreadyConsecutive = Txt("Die ") & claimCount & _
        Txt(" Anspr{ue}che sind bereits fortlaufend nummeriert, alle R{ue}ckbez{ue}ge passen.")
End Function

Public Function MsgRenumberDone(ByVal claimCount As Long, ByVal labelsChanged As Long, _
        ByVal refsChanged As Long, ByVal flagCount As Long, ByVal commentAuthor As String) As String
    Dim s As String
    s = Txt("Neunummerierung abgeschlossen.") & vbCr & vbCr & _
        Txt("Anspr{ue}che: ") & claimCount & vbCr & _
        Txt("Ge{ae}nderte Anspruchsnummern: ") & labelsChanged & vbCr & _
        Txt("Aktualisierte R{ue}ckbez{ue}ge: ") & refsChanged
    If flagCount > 0 Then
        s = s & vbCr & vbCr & Txt("Bitte pr{ue}fen: ") & flagCount & _
            Txt(" Hinweis(e) als Kommentar von ") & Quoted(commentAuthor) & Txt(" eingef{ue}gt.")
    End If
    MsgRenumberDone = s
End Function

Public Function CmtUnknownReference(ByVal label As String) As String
    CmtUnknownReference = Txt("R{ue}ckbezug auf nicht vorhandenen Anspruch ") & Quoted(label) & _
        Txt(" {-} nicht ge{ae}ndert. Bitte pr{ue}fen.")
End Function

Public Function CmtRangeCheck(ByVal oldFrom As String, ByVal oldTo As String, _
        ByVal newFrom As Long, ByVal newTo As Long) As String
    If newFrom > newTo Then
        CmtRangeCheck = Txt("Bereich ") & oldFrom & " bis " & oldTo & _
            Txt(" ist nach der Neunummerierung umgekehrt (") & newFrom & " bis " & newTo & _
            Txt("). Bitte pr{ue}fen.")
    Else
        CmtRangeCheck = Txt("Bereich bitte pr{ue}fen: vorher ") & oldFrom & " bis " & oldTo & _
            ", jetzt " & newFrom & " bis " & newTo & _
            Txt(". Der Bereich umfasst die Anspr{ue}che in Dokumentreihenfolge; ") & _
            Txt("eingef{ue}gte oder verschobene Anspr{ue}che sind darin enthalten bzw. nicht mehr enthalten.")
    End If
End Function

Public Function CmtManualCheck(ByVal oldText As String, ByVal newText As String) As String
    CmtManualCheck = Txt("Nummer ") & Quoted(oldText) & Txt(" konnte nicht automatisch in ") & Quoted(newText) & _
        Txt(" ge{ae}ndert werden (Feld oder {Ae}nderungsmarkierung im Text). Bitte manuell anpassen.")
End Function
