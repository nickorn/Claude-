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

Public Const APP_VERSION As String = "0.6.0"

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

' ---------------------------------------------------------------- claim check (MVP-2)

Public Function CheckDisplayName(ByVal checkKey As String) As String
    Select Case checkKey
        Case "numbering": CheckDisplayName = "Nummerierung"
        Case "punctuation": CheckDisplayName = "Zeichensetzung"
        Case "spacing": CheckDisplayName = Txt("Abst{ae}nde")
        Case "formatting": CheckDisplayName = "Formatierung"
        Case "backReferences": CheckDisplayName = Txt("R{ue}ckbez{ue}ge")
        Case "independentClaims": CheckDisplayName = Txt("Unabh{ae}ngige Anspr{ue}che")
        Case "multipleDependencies": CheckDisplayName = Txt("Mehrfachabh{ae}ngigkeit")
        Case "antecedentBasis": CheckDisplayName = "Bezugsbasis"
        Case "referenceSigns": CheckDisplayName = "Bezugszeichen"
        Case "blacklist": CheckDisplayName = Txt("Unklare Begriffe")
        Case "categories": CheckDisplayName = "Kategorien"
        Case "renumber": CheckDisplayName = "Neu nummerieren"
        Case Else: CheckDisplayName = checkKey
    End Select
End Function

Public Function CategoryName(ByVal cat As String) As String
    Select Case cat
        Case "method": CategoryName = "Verfahren"
        Case "device": CategoryName = "Vorrichtung"
        Case "system": CategoryName = "System"
        Case "program": CategoryName = "Computerprogramm"
        Case "medium": CategoryName = "Speichermedium"
        Case "use": CategoryName = "Verwendung"
        Case Else: CategoryName = "unbekannt"
    End Select
End Function

Public Function UndoTitleCheck() As String
    UndoTitleCheck = AppTitle() & Txt(": Pr{ue}fung")
End Function

Public Function UndoTitleRemoveMarks() As String
    UndoTitleRemoveMarks = AppTitle() & Txt(": Pr{ue}fmarkierungen entfernen")
End Function

Public Function MsgNumDuplicate(ByVal label As String) As String
    MsgNumDuplicate = Txt("Anspruchsnummer ") & Quoted(label) & Txt(" kommt mehrfach vor.")
End Function

Public Function MsgNumSupplementary(ByVal label As String) As String
    MsgNumSupplementary = Txt("Zusatznummer ") & Quoted(label) & Txt(" {-} bitte {q<}Neu nummerieren{q>} ausf{ue}hren.")
End Function

Public Function MsgNumWrong(ByVal label As String, ByVal expected As Long) As String
    MsgNumWrong = Txt("Nummer ") & Quoted(label) & Txt(" an dieser Stelle; erwartet: ") & expected & _
        Txt(" (L{ue}cke oder falsche Reihenfolge).")
End Function

Public Function MsgPunctNoFullStop() As String
    MsgPunctNoFullStop = Txt("Anspruch endet nicht mit einem Punkt.")
End Function

Public Function MsgPunctDoubleFullStop() As String
    MsgPunctDoubleFullStop = Txt("Anspruch endet mit mehr als einem Punkt.")
End Function

Public Function MsgPunctFullStopInside() As String
    MsgPunctFullStopInside = Txt("Punkt vor dem Ende des Anspruchs {-} danach folgt noch Text (z. B. {q<}wherein{q>}/{q<}wobei{q>}).")
End Function

Public Function MsgPunctAndMisplaced() As String
    MsgPunctAndMisplaced = Txt("{q<}; and{q>}/{q<}; und{q>} steht nicht vor dem letzten Merkmal der Aufz{ae}hlung.")
End Function

Public Function MsgPunctFeatureEnd() As String
    MsgPunctFeatureEnd = Txt("Merkmal endet nicht mit {q<};{q>} (oder {q<}:{q>} vor einer Aufz{ae}hlung).")
End Function

Public Function MsgPunctAndMissing(ByVal lang As String) As String
    If lang = "de" Then
        MsgPunctAndMissing = Txt("Vor dem letzten Merkmal fehlt {q<}; und{q>}.")
    Else
        MsgPunctAndMissing = Txt("Vor dem letzten Merkmal fehlt {q<}; and{q>}.")
    End If
End Function

Public Function MsgSpacingBetweenClaims(ByVal prevLabel As String, ByVal label As String, ByVal gap As Long) As String
    MsgSpacingBetweenClaims = Txt("Zwischen Anspruch ") & prevLabel & " und " & label & " stehen " & gap & _
        Txt(" Leerabs{ae}tze (erwartet: genau 1).")
End Function

Public Function MsgSpacingDouble() As String
    MsgSpacingDouble = Txt("Doppeltes Leerzeichen.")
End Function

Public Function MsgSpacingBeforePunct(ByVal mark As String) As String
    MsgSpacingBeforePunct = Txt("Leerzeichen vor ") & Quoted(mark) & "."
End Function

Public Function MsgRefUnknown(ByVal label As String) As String
    MsgRefUnknown = Txt("R{ue}ckbezug auf nicht vorhandenen Anspruch ") & Quoted(label) & "."
End Function

Public Function MsgRefSelf() As String
    MsgRefSelf = Txt("Anspruch bezieht sich auf sich selbst.")
End Function

Public Function MsgRefForward(ByVal label As String) As String
    MsgRefForward = Txt("R{ue}ckbezug auf den sp{ae}teren Anspruch ") & Quoted(label) & Txt(" (Vorw{ae}rtsbezug).")
End Function

Public Function MsgIndependentWithRef(ByVal label As String) As String
    MsgIndependentWithRef = Txt("Unabh{ae}ngiger Anspruch mit R{ue}ckbezug auf Anspruch ") & Quoted(label) & _
        Txt(" derselben Kategorie.")
End Function

Public Function MsgMultiDep() As String
    MsgMultiDep = Txt("Mehrfach abh{ae}ngiger Anspruch (US: Zusatzgeb{ue}hr).")
End Function

Public Function MsgMultiOnMulti(ByVal label As String) As String
    MsgMultiOnMulti = Txt("Mehrfach abh{ae}ngiger Anspruch bezieht sich auf den mehrfach abh{ae}ngigen Anspruch ") & _
        Quoted(label) & Txt(" (US: unzul{ae}ssig).")
End Function

Public Function MsgCategoryMismatch(ByVal cat As String, ByVal label As String, ByVal targetCat As String) As String
    MsgCategoryMismatch = Txt("Kategorie passt nicht: ") & CategoryName(cat) & Txt(" bezieht sich auf Anspruch ") & _
        Quoted(label) & " (" & CategoryName(targetCat) & ")."
End Function

Public Function MsgCategorySecondIndependent(ByVal cat As String, ByVal label As String) As String
    MsgCategorySecondIndependent = Txt("Weiterer unabh{ae}ngiger Anspruch der Kategorie ") & CategoryName(cat) & _
        Txt(" (vgl. Anspruch ") & label & Txt("; Regel 43(2) EP{Ue}).")
End Function

Public Function MsgAntecedentMissing(ByVal phrase As String) As String
    MsgAntecedentMissing = Txt("Keine Bezugsbasis f{ue}r ") & Quoted(phrase) & _
        Txt(" {-} Begriff wurde vorher nicht eingef{ue}hrt (a/an/ein/eine).")
End Function

Public Function MsgAntecedentPlural(ByVal phrase As String) As String
    MsgAntecedentPlural = Txt("Singular/Plural passt nicht zur Einf{ue}hrung: ") & Quoted(phrase) & "."
End Function

Public Function MsgSignTermDiffers(ByVal phrase As String, ByVal signs As String, ByVal previous As String) As String
    MsgSignTermDiffers = Txt("Gleicher Begriff ") & Quoted(phrase) & Txt(" mit anderem Bezugszeichen (") & signs & _
        Txt(", vorher ") & previous & ")."
End Function

Public Function MsgSignUsedTwice(ByVal sign As String, ByVal otherTerm As String) As String
    MsgSignUsedTwice = Txt("Bezugszeichen ") & sign & Txt(" steht bereits f{ue}r ") & Quoted(otherTerm) & "."
End Function

Public Function MsgSignMissing(ByVal phrase As String, ByVal signs As String) As String
    MsgSignMissing = Txt("Bezugszeichen fehlt f{ue}r ") & Quoted(phrase) & " (" & signs & ")."
End Function

Public Function MsgSignSemicolon(ByVal separator As String) As String
    MsgSignSemicolon = Txt("Bezugszeichen mit {q<};{q>} getrennt {-} Trennzeichen ") & Quoted(separator) & _
        Txt(" verwenden.")
End Function

Public Function MsgSignSeparator(ByVal separator As String) As String
    MsgSignSeparator = Txt("Bezugszeichen bitte mit ") & Quoted(separator) & Txt(" trennen.")
End Function

Public Function MsgBlacklist(ByVal term As String) As String
    MsgBlacklist = Txt("M{oe}glicherweise unklarer Begriff ") & Quoted(term) & Txt(" (Art. 84 EP{Ue} / {ss} 34 PatG).")
End Function

Public Function MsgFormatParagraph(ByVal details As String) As String
    MsgFormatParagraph = Txt("Formatierung weicht ab: ") & details & "."
End Function

Public Function MsgFormatFont(ByVal actual As String, ByVal expected As String) As String
    If Len(actual) = 0 Then actual = "gemischt"
    MsgFormatFont = "Schriftart " & actual & " statt " & expected
End Function

Public Function MsgFormatSize(ByVal actual As String, ByVal expected As String) As String
    MsgFormatSize = Txt("Schriftgr{oe}{ss}e ") & actual & " statt " & expected & " pt"
End Function

Public Function MsgFormatLineSpacing(ByVal expected As String) As String
    MsgFormatLineSpacing = "Zeilenabstand nicht " & expected & "-fach"
End Function

Public Function MsgFormatParaSpacing() As String
    MsgFormatParaSpacing = Txt("Abstand vor/nach nicht 0 pt")
End Function

Public Function MsgFormatMore(ByVal count As Long) As String
    MsgFormatMore = Txt("Weitere ") & count & Txt(" Abs{ae}tze mit Formatabweichungen {-} ") & _
        Txt("{q<}Formatierung anwenden{q>} korrigiert alle auf einmal.")
End Function

Public Function MsgCheckSummary(ByVal lines As String, ByVal total As Long, ByVal marked As Boolean) As String
    If total = 0 Then
        MsgCheckSummary = Txt("Pr{ue}fung abgeschlossen {-} keine Auff{ae}lligkeiten gefunden.")
        Exit Function
    End If
    MsgCheckSummary = Txt("Pr{ue}fung abgeschlossen: ") & total & Txt(" Hinweis(e).") & vbCr & vbCr & lines
    If marked Then
        MsgCheckSummary = MsgCheckSummary & vbCr & Txt("Die Stellen sind kommentiert und farbig markiert. ") & _
            Txt("Entfernen mit {q<}Pr{ue}fmarkierungen entfernen{q>}.")
    End If
End Function

Public Function MsgMarksRemoved(ByVal count As Long) As String
    MsgMarksRemoved = count & Txt(" Pr{ue}fmarkierung(en) entfernt.")
End Function

Public Function ReportTitle(ByVal docName As String) As String
    ReportTitle = Txt("Pr{ue}fbericht {-} ") & docName & " (" & Format$(Now, "dd.mm.yyyy hh:nn") & ")"
End Function

Public Function ReportHeader() As String
    ReportHeader = "Anspruch" & vbTab & Txt("Pr{ue}fung") & vbTab & "Befund" & vbTab & "Textauszug"
End Function

Public Function ReportNoFindings() As String
    ReportNoFindings = Txt("Keine Auff{ae}lligkeiten gefunden.")
End Function

' ---------------------------------------------------------------- formatting (MVP-3)

' kind: 0 = layout, 1 = spaces, 2 = punctuation, 3 = reference signs
Public Function FormatCommandName(ByVal kind As Long) As String
    Select Case kind
        Case 0: FormatCommandName = "Formatierung anwenden"
        Case 1: FormatCommandName = "Doppelte Leerzeichen"
        Case 2: FormatCommandName = "Zeichensetzung"
        Case 3: FormatCommandName = "Bezugszeichen-Format"
    End Select
End Function

Public Function UndoTitleFormat(ByVal kind As Long) As String
    UndoTitleFormat = AppTitle() & ": " & FormatCommandName(kind)
End Function

Public Function MsgFormatTextDone(ByVal kind As Long, ByVal applied As Long, ByVal failed As Long) As String
    Dim s As String
    If applied = 0 And failed = 0 Then
        MsgFormatTextDone = FormatCommandName(kind) & Txt(": nichts zu korrigieren.")
        Exit Function
    End If
    s = FormatCommandName(kind) & ": " & applied & Txt(" Stelle(n) korrigiert.")
    If failed > 0 Then s = s & vbCr & failed & _
        Txt(" Stelle(n) konnten wegen Feldern oder {Ae}nderungsmarkierungen nicht ge{ae}ndert werden {-} bitte Vollpr{ue}fung ausf{ue}hren.")
    If kind = 2 Then s = s & vbCr & vbCr & _
        Txt("Nicht automatisch ge{ae}ndert: falsch platziertes {q<}; and{q>} und Text nach dem Schlusspunkt (siehe Vollpr{ue}fung).")
    MsgFormatTextDone = s
End Function

Public Function MsgFormatLayoutDone(ByVal added As Long, ByVal removed As Long, ByVal indented As Long) As String
    MsgFormatLayoutDone = Txt("Formatierung angewendet (Schrift, Zeilen- und Absatzabstand).") & vbCr & vbCr & _
        Txt("Leerabs{ae}tze eingef{ue}gt: ") & added & vbCr & _
        Txt("Leerabs{ae}tze entfernt: ") & removed & vbCr & _
        Txt("Einz{ue}ge angepasst: ") & indented
End Function

' ---------------------------------------------------------------- versions + footer (roadmap 4)

Public Function UndoTitleFooter() As String
    UndoTitleFooter = AppTitle() & Txt(": Fu{ss}zeile/Status")
End Function

Public Function MsgVersionSaveFirst() As String
    MsgVersionSaveFirst = Txt("Bitte das Dokument zuerst speichern. Die neue Fassung wird aus der gespeicherten Datei erzeugt; ") & _
        Txt("das Original bleibt unver{ae}ndert.")
End Function

Public Function MsgVersionAskSave(ByVal docName As String) As String
    MsgVersionAskSave = Quoted(docName) & Txt(" hat ungespeicherte {Ae}nderungen. Jetzt speichern und die neue Fassung erzeugen?") & _
        vbCr & vbCr & Txt("(Das Original wird nur gespeichert, nicht ver{ae}ndert.)")
End Function

Public Function MsgVersionCreated(ByVal status As String) As String
    MsgVersionCreated = Txt("Neue Fassung erzeugt: ") & Quoted(status) & vbCr & vbCr & _
        Txt("Das Dokument ist noch nicht gespeichert {-} bitte unter neuem Namen speichern. Das Original ist unver{ae}ndert.")
End Function

Public Function MsgVersionNoRevisions() As String
    MsgVersionNoRevisions = Txt("Markierte Fassung erzeugt {-} das Dokument enth{ae}lt jedoch keine nachverfolgten {Ae}nderungen.") & _
        vbCr & vbCr & Txt("Bitte unter neuem Namen speichern. Das Original ist unver{ae}ndert.")
End Function

Public Function MsgVersionLostRevisions() As String
    MsgVersionLostRevisions = Txt("Achtung: Die nachverfolgten {Ae}nderungen wurden nicht in die Kopie {ue}bernommen. ") & _
        Txt("Bitte die markierte Fassung manuell pr{ue}fen (das Original ist unver{ae}ndert) und diesen Fall melden.")
End Function

Public Function MsgFooterSet(ByVal status As String, ByVal count As Long) As String
    If Len(status) = 0 Then
        MsgFooterSet = Txt("Status in der Fu{ss}zeile entfernt (") & count & Txt(" Fu{ss}zeile(n)).")
    Else
        MsgFooterSet = Txt("Fu{ss}zeile links: ") & Quoted(status) & " (" & count & Txt(" Fu{ss}zeile(n)). ") & _
            Txt("Die rechte Seite (Anmeldenummer) bleibt unver{ae}ndert.")
    End If
End Function

' ---------------------------------------------------------------- text blocks (roadmap 5)

Public Function UndoTitleTextBlock() As String
    UndoTitleTextBlock = AppTitle() & Txt(": Textbaustein einf{ue}gen")
End Function

Public Function MsgTextBlockMissing(ByVal path As String) As String
    MsgTextBlockMissing = Txt("Der Textbaustein wurde nicht gefunden:") & vbCr & path
End Function

Public Function MsgPlaceholderPrompt(ByVal name As String) As String
    MsgPlaceholderPrompt = Txt("Wert f{ue}r den Platzhalter [") & name & "]:" & vbCr & vbCr & _
        Txt("(Leer lassen oder Abbrechen: Platzhalter bleibt stehen und wird gelb markiert.)")
End Function

Public Function MsgTextBlocksImported(ByVal count As Long, ByVal folder As String) As String
    MsgTextBlocksImported = count & Txt(" Textbaustein(e) importiert nach:") & vbCr & folder & vbCr & vbCr & _
        Txt("Vorhandene Dateien wurden nicht {ue}berschrieben.")
End Function
