Attribute VB_Name = "mdl_Translation"
Option Explicit
Option Private Module

Public strLabel(4) As String
Public strScreentip(4) As String
Public strSupertip(4) As String
Public strError(5) As String
Public strCmd(4) As String
Public strFrmInfo(1) As String
Public strFrmHolidays(7) As String
Public strFrmFunction(8) As String

Public Sub germanText()
    strFrmInfo(0) = "Der Quellcode is OpenSource unter AGPLv3 und verfügbar auf "
    
    strCmd(0) = "OK"
    strCmd(1) = "Abbrechen"
    strCmd(2) = "Auflisten"
    strCmd(3) = "Wert einfügen"
    strCmd(4) = "Funktion einfügen"
    
    strFrmHolidays(0) = "Feiertage importieren"
    strFrmHolidays(1) = "Für Jahr:"
    strFrmHolidays(2) = "Importieren in Zelle:"
    strFrmHolidays(3) = "Für Land:"
    strFrmHolidays(4) = "Es müssen ab der ausgewählten Zelle 3 Spalten und  20 Zeilen Platz sein. Gefüllte Zellen werden ohne Rückfrage überschrieben."
    strFrmHolidays(5) = "Das GivenYear muss als Zahl angegeben sein."
    strFrmHolidays(6) = "Es muss eine Zelle ausgewählt werden."
    strFrmHolidays(7) = "Es darf nur eine Zelle ausgewählt werden."

    strFrmFunction(0) = "Jahr (GivenYear):"
    strFrmFunction(1) = "Easter-Funktion"
    strFrmFunction(2) = "Die Funktion Easter(GivenYear) gibt das Datum des Ostersonntags für das gegebene Jahr zurück."
    strFrmFunction(3) = "LastAdvent-Funktion"
    strFrmFunction(4) = "Die Funktion LastAdvent(GivenYear) gibt das Datum des 4. Adventsonntags für das gegebene Jahr zurück."
    strFrmFunction(5) = "dd.MM.yyyy"
    strFrmFunction(6) = "Bitte nur eine Zelle auswählen!"
    strFrmFunction(7) = "Die Zelle muss eine Zahl enthalten."
    strFrmFunction(8) = "Es wurde kein gültiger Bereich eingegeben."

    strError(0) = "Fehler"
    
    strLabel(0) = "Feiertage importieren"
    strLabel(1) = "Funktion Easter"
    strLabel(2) = "Funktion LastAdvent"
    strLabel(3) = ""
    strLabel(4) = "Info"
    

    strScreentip(0) = "Importiert die Feiertage eines gegebenen Jahres."
    strScreentip(1) = "Funktion Easter gibt das Datum des Ostersonntags eines gegebenen Jahres zurück."
    strScreentip(2) = "Funktion LastAdvent gibt das Datum des 4. Adventsonntags eines gegebenen Jahres zurück."
    strScreentip(3) = ""
    strScreentip(4) = ""
    
    strSupertip(0) = "Importiert die Feiertage eines gegebenen Jahres."
    strSupertip(1) = "Funktion Easter gibt das Datum des Ostersonntags eines gegebenen Jahres zurück."
    strSupertip(2) = "Funktion LastAdvent gibt das Datum des 4. Adventsonntags eines gegebenen Jahres zurück."
    strSupertip(3) = ""
    strSupertip(4) = ""
End Sub

Public Sub englishText()
    strFrmInfo(0) = "Source code is OpenSource under AGPLv3 and available at "
    
    strCmd(0) = "OK"
    strCmd(1) = "Cancel"
    strCmd(2) = "List"
    strCmd(3) = "Insert value"
    strCmd(4) = "Insert function"
    
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
    strFrmFunction(8) = "Es wurde kein gültiger Bereich eingegeben."

    strError(0) = "Fehler"
    
    strLabel(0) = "Feiertage importieren"
    strLabel(1) = "Funktion Easter"
    strLabel(2) = "Funktion LastAdvent"
    strLabel(3) = ""
    strLabel(4) = "Info"

End Sub

