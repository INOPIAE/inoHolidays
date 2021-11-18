Attribute VB_Name = "mdl_Translation"
Option Explicit
Option Private Module

Public strLabel(5) As String
Public strScreentip(5) As String
Public strSupertip(5) As String
Public strError(5) As String
Public strCmd(7) As String
Public strFrmInfo(1) As String
Public strFrmHolidays(7) As String
Public strFrmFunction(8) As String
Public strRegister(7) As String
Public strFrmOutlook(12) As String
Public strCountry(3) As String

Public Sub germanText()
    strFrmInfo(0) = "Der Quellcode ist OpenSource unter AGPLv3." & vbNewLine & "Der Quellcode und die Dokumentation sind verf�gbar auf "
    
    strCmd(0) = "OK"
    strCmd(1) = "Abbrechen"
    strCmd(2) = "Auflisten"
    strCmd(3) = "Wert einf�gen"
    strCmd(4) = "Funktion einf�gen"
    strCmd(5) = "Importieren"
    strCmd(6) = "Jahr l�schen"
    strCmd(7) = "Alle l�schen"
    
    strFrmHolidays(0) = "Feiertage importieren"
    strFrmHolidays(1) = "F�r Jahr:"
    strFrmHolidays(2) = "Importieren in Zelle:"
    strFrmHolidays(3) = "F�r Land:"
    strFrmHolidays(4) = "Es m�ssen ab der ausgew�hlten Zelle 3 Spalten und  20 Zeilen Platz sein. Gef�llte Zellen werden ohne R�ckfrage �berschrieben."
    strFrmHolidays(5) = "Das Jahr muss als Zahl angegeben sein."
    strFrmHolidays(6) = "Es muss eine Zelle ausgew�hlt werden."
    strFrmHolidays(7) = "Es darf nur eine Zelle ausgew�hlt werden."

    strFrmFunction(0) = "Jahr (GivenYear):"
    strFrmFunction(1) = "Easter-Funktion"
    strFrmFunction(2) = "Die Funktion Easter(GivenYear) gibt das Datum des Ostersonntags f�r das gegebene Jahr zur�ck."
    strFrmFunction(3) = "LastAdvent-Funktion"
    strFrmFunction(4) = "Die Funktion LastAdvent(GivenYear) gibt das Datum des 4. Adventsonntags f�r das gegebene Jahr zur�ck."
    strFrmFunction(5) = "dd.MM.yyyy"
    strFrmFunction(6) = "Bitte nur eine Zelle ausw�hlen!"
    strFrmFunction(7) = "Die Zelle muss eine Zahl enthalten."
    strFrmFunction(8) = "Es wurde kein g�ltiger Bereich eingegeben."

    strError(0) = "Fehler"
    
    strLabel(0) = "Feiertage importieren"
    strLabel(1) = "Funktion Easter"
    strLabel(2) = "Funktion LastAdvent"
    strLabel(3) = ""
    strLabel(4) = "Info"
    strLabel(5) = "Feiertage in Outlook importieren"

    strScreentip(0) = "Importiert die Feiertage eines gegebenen Jahres."
    strScreentip(1) = "Funktion Easter gibt das Datum des Ostersonntags eines gegebenen Jahres zur�ck."
    strScreentip(2) = "Funktion LastAdvent gibt das Datum des 4. Adventsonntags eines gegebenen Jahres zur�ck."
    strScreentip(3) = ""
    strScreentip(4) = ""
    strScreentip(5) = "Importieren der Feiertage nach Outlook"
    
    strSupertip(0) = "Importiert die Feiertage eines gegebenen Jahres."
    strSupertip(1) = "Funktion Easter gibt das Datum des Ostersonntags eines gegebenen Jahres zur�ck."
    strSupertip(2) = "Funktion LastAdvent gibt das Datum des 4. Adventsonntags eines gegebenen Jahres zur�ck."
    strSupertip(3) = ""
    strSupertip(4) = ""
    strSupertip(5) = "Importieren der Feiertage nach Outlook"
    
    strRegister(0) = "GivenYear - Das Jahr f�r den Ostersonntag an"
    strRegister(1) = "Gibt das Datum des Ostersonntags des angegeben Jahres (GivenYear) zur�ck."
    strRegister(2) = "GivenYear - Das Jahr f�r den 4. Adventsonntag an"
    strRegister(3) = "Gibt das Datum des 4. Adventsonntags des angegeben Jahres (GivenYear) zur�ck."
    strRegister(4) = "GivenDate - Das Datum das �berpr�ft werden soll"
    strRegister(5) = "Country - Der Staat (Country) in 2-Zeichen-ISO-Code f�r den der Feiertag ermittelt werden soll." _
        & "Standardvorgabe ist 'de'."
    strRegister(6) = "State - Das Bundesland f�r das der Feiertag ermittelt werden soll. (siehe Dokumenation)" _
        & "Es gibt keine Standardvorgabe."
    strRegister(7) = "Pr�ft, ob das angegebene Datum (GivenDate) unter Ber�cksichtigung des Staates und evtl. Bundeslandes ein Feiertag ist."

    strFrmOutlook(0) = "Feiertage nach Outlook importieren"
    strFrmOutlook(1) = "F�r Jahr:"
    strFrmOutlook(2) = "F�r Land:"
    strFrmOutlook(3) = "F�r Bundesland/Region:"
    strFrmOutlook(4) = "Feiertage als gebucht eintragen"
    strFrmOutlook(5) = "Die ersten 3 Felder m�ssen gef�llt sein." & vbNewLine & "Wenn 'Feiertage als gebucht eintragen' ausgew�hlt ist, werden die Feiertage f�r das ausgew�hlte Bundesland/Region als gebucht eingetragen. Bundesland/Region =  All nur die landesweiten Feiertage werden eingetragen."
    strFrmOutlook(6) = "Das Jahr muss als Zahl angegeben sein."
    strFrmOutlook(7) = "Es muss ein Land ausgew�hlt sein."
    strFrmOutlook(8) = "Es muss ein Bundesland ausgew�hlt sein."
    strFrmOutlook(9) = "Outlook muss ge�ffnet sein."
    strFrmOutlook(10) = "{0} Eintr�ge f�r {1} bearbeitet. Davon " & vbNewLine _
        & "{2} neue Eintr�ge" & vbNewLine _
        & "{3} ge�nderte Eintr�ge" & vbNewLine _
        & "{4} unver�nderte Eintr�ge" & vbNewLine
    strFrmOutlook(11) = "{0} Eintr�ge f�r {1} gel�scht."
    strFrmOutlook(12) = "{0} Eintr�ge gel�scht."

    strCountry(0) = "Deutschland"
    strCountry(1) = "�sterreich"
    strCountry(2) = "Schweiz"
    strCountry(3) = "Italien"
End Sub

Public Sub englishText()
    strFrmInfo(0) = "The source code is OpenSource under AGPLv3." & vbNewLine & "The source code and the documentation are available at "
    
    strCmd(0) = "OK"
    strCmd(1) = "Cancel"
    strCmd(2) = "List"
    strCmd(3) = "Insert value"
    strCmd(4) = "Insert function"
    strCmd(5) = "Import"
    strCmd(6) = "Delete year"
    strCmd(7) = "Delete all"
    
    strFrmHolidays(0) = "Import Holidays"
    strFrmHolidays(1) = "Year:"
    strFrmHolidays(2) = "Import into cell:"
    strFrmHolidays(3) = "Country:"
    strFrmHolidays(4) = "There should be 3 emtpy columns and 20 empty rows starting from the given cell. Filled cell will be overwriten without further notification."
    strFrmHolidays(5) = "GivenYear needs to be entered as number."
    strFrmHolidays(6) = "Select a cell."
    strFrmHolidays(7) = "Only one cell must be selected."

    strFrmFunction(0) = "GivenYear:"
    strFrmFunction(1) = "Function Easter"
    strFrmFunction(2) = "The function Easter(GivenYear) returns the date of Easter Sunday of the given year."
    strFrmFunction(3) = "Function LastAdvent"
    strFrmFunction(4) = "The function LastAdvent(GivenYear) returns the date of 4th Advent Sunday of the given year."
    strFrmFunction(5) = "MM/dd/yyyy"
    strFrmFunction(6) = "Only one cell must be selected."
    strFrmFunction(7) = "The cell must contain a number."
    strFrmFunction(8) = "No valid range entered."

    strError(0) = "Error"
    
    strLabel(0) = "Import Holidays"
    strLabel(1) = "Function Easter"
    strLabel(2) = "Function LastAdvent"
    strLabel(3) = ""
    strLabel(4) = "Info"
    strLabel(5) = "Add Holidays to Outlook"
    
    strScreentip(0) = "Import holidays of a given Year."
    strScreentip(1) = "Function Easter returns the date of Easter sunday of a given year."
    strScreentip(2) = "Function LastAdvent returns the date of the 4th Sunday in Advent of a given year."
    strScreentip(3) = ""
    strScreentip(4) = ""
    strScreentip(5) = "Add Holidays to Outlook"
    
    strSupertip(0) = "Import holidays of a given Year."
    strSupertip(1) = "Function Easter returns the date of Easter sunday of a given year."
    strSupertip(2) = "Function LastAdvent returns the date of the 4th Sunday in Advent of a given year."
    strSupertip(3) = ""
    strSupertip(4) = ""
    strSupertip(5) = "Add Holidays to Outlook"
        
    strRegister(0) = "GivenYear - Year for the Easter Sunday"
    strRegister(1) = "Returns the date of Easter Sunday of the given year"
    strRegister(2) = "GivenYear - Year for the 4th Advent Sunday"
    strRegister(3) = "Returns the date of 4th Advent Sunday of the given year"
    strRegister(4) = "GivenDate - Date to be checked"
    strRegister(5) = "Country - Country in 2-letter-ISO-Code for which the holiday shall be checked." _
        & "Default value is 'de'."
    strRegister(6) = "State - State for which the holiday shall be checked. (see documantation)" _
        & "No default value given."
    strRegister(7) = "Checks whether the given date is a holiday for a given country and tentative state."


    strFrmOutlook(0) = "Add Holidays To Outlook"
    strFrmOutlook(1) = "Year:"
    strFrmOutlook(2) = "Country:"
    strFrmOutlook(3) = "State/Region:"
    strFrmOutlook(4) = "Add holiday as busy"
    strFrmOutlook(5) = "The first 3 fields must be used." & vbNewLine & "If 'Add holiday as busy' is ticked the public holidays will be add as busy for the given State/Region. If State/Region = All only the countrywide holidays are set as busy."
    strFrmOutlook(6) = "Year must be entered as number."
    strFrmOutlook(7) = "A country must be selected."
    strFrmOutlook(8) = "A state must be selected."
    strFrmOutlook(9) = "Outlook must be started."
    strFrmOutlook(10) = "{0} entries processed for {1}. Thereof " & vbNewLine _
        & "{2} new entries" & vbNewLine _
        & "{3} changed entries" & vbNewLine _
        & "{4} unchanged entries" & vbNewLine
    strFrmOutlook(11) = "{0} entries deleted for {1}."
    strFrmOutlook(12) = "{0} entries deleted."

    strCountry(0) = "Germany"
    strCountry(1) = "Austria"
    strCountry(2) = "Switzerland"
    strCountry(3) = "Italy"
End Sub

