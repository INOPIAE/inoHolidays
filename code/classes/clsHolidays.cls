VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHolidays"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'https://www.datenportal.bmbf.de/portal/de/G122.html

Public Jahr As Integer
Public Country As String
Private Countries() As String

Private Type Holiday
    Datum As Date
    Holiday As String
    State() As Variant
End Type

Private Holidays() As Holiday

Private Sub Class_Initialize()
    Jahr = Year(Now)
    Country = "de"
    InitDatumFile
    FillCountries
End Sub

Private Sub InitDatum()
    ReDim Holidays(16)
    Holidays(0).Datum = DateSerial(Jahr, 1, 1)
    Holidays(0).Holiday = "Neujahr"
    Holidays(0).State = Array("All")
    Holidays(1).Datum = DateSerial(Jahr, 1, 6)
    Holidays(1).Holiday = "Drei Hl. Könige"
    Holidays(1).State = Array("BY", "ST", "BW")
    Holidays(2).Datum = DateSerial(Jahr, 3, 8)
    Holidays(2).Holiday = "Int. Frauentag"
    Holidays(2).State = Array("BE")
    Dim Ostersonntag As Date
    Ostersonntag = Ostern(Jahr)
    Holidays(3).Datum = DateAdd("d", -2, Ostersonntag)
    Holidays(3).Holiday = "Karfreitag"
    Holidays(3).State = Array("All")
    Holidays(4).Datum = Ostersonntag
    Holidays(4).Holiday = "Ostersonntag"
    Holidays(4).State = Array("All")
    Holidays(5).Datum = DateAdd("d", 1, Ostersonntag)
    Holidays(5).Holiday = "Ostermontag"
    Holidays(5).State = Array("All")
    Holidays(6).Datum = DateAdd("d", 39, Ostersonntag)
    Holidays(6).Holiday = "Christi Himmelfahrt"
    Holidays(6).State = Array("All")
    Holidays(7).Datum = DateAdd("d", 50, Ostersonntag)
    Holidays(7).Holiday = "Pfingstmontag"
    Holidays(7).State = Array("All")
    Holidays(8).Datum = DateAdd("d", 60, Ostersonntag)
    Holidays(8).Holiday = "Fronleichnam"
    Holidays(8).State = Array("BW", "BY", "HE", "NW", "RP", "SL")
    Holidays(9).Datum = DateSerial(Jahr, 8, 15)
    Holidays(9).Holiday = "Maria Himmelfahrt"
    Holidays(9).State = Array("BY", "SL")
    Holidays(10).Datum = DateSerial(Jahr, 9, 20)
    Holidays(10).Holiday = "Weltkindertag"
    Holidays(10).State = Array("TH")
    Holidays(11).Datum = DateSerial(Jahr, 10, 3)
    Holidays(11).Holiday = "Tag der dt. Einheit"
    Holidays(11).State = Array("All")
    Holidays(12).Datum = DateSerial(Jahr, 10, 31)
    Holidays(12).Holiday = "Reformationstag"
    Holidays(12).State = Array("SH", "NI", "HB", "HH", "BB", "ST", "SN", "TH", "MV")
    Holidays(13).Datum = DateSerial(Jahr, 11, 1)
    Holidays(13).Holiday = "Allerheiligen"
    Holidays(13).State = Array("BW", "BY", "NW", "RP", "SL")
    Dim LetzterAdvent As Date
    LetzterAdvent = LetzterAdventSonntag(Jahr)
    Holidays(14).Datum = DateAdd("d", -32, LetzterAdvent)
    Holidays(14).Holiday = "Buß- und Bettag"
    Holidays(14).State = Array("SN")
    Holidays(15).Datum = DateSerial(Jahr, 12, 25)
    Holidays(15).Holiday = "1. Weihnachtstag"
    Holidays(15).State = Array("All")
    Holidays(16).Datum = DateSerial(Jahr, 12, 26)
    Holidays(16).Holiday = "2. Weihnachtstag"
    Holidays(16).State = Array("All")

End Sub

Private Sub InitDatumFile()
    Dim FNr As Integer
    Dim intLine As Integer
    Dim DataLine As String
    Dim Ostersonntag As Date
    Ostersonntag = Ostern(Jahr)
    Dim LetzterAdvent As Date
    LetzterAdvent = LetzterAdventSonntag(Jahr)

    Dim myLine() As String
    FNr = FreeFile
    Dim strFile As String
    strFile = AddIns(strVBProjects).Path & "\countrycodes\" & Country & ".inocsv"
    
    If mdlHelper_FileExists(strFile) = False Then
        InitDatum
        Exit Sub
    End If
    
    Open strFile For Input As FNr
    Do Until EOF(FNr)
        ReDim Preserve Holidays(intLine)
        Line Input #FNr, DataLine
        myLine = Split(DataLine, ";")
        If IsDate(myLine(0)) Then
            Holidays(intLine).Datum = DateSerial(Jahr, Month(myLine(0)), Day(myLine(0)))
        Else
            Dim pos As Integer
            Dim intSign As Integer
            intSign = 1
            pos = InStr(myLine(0), "-")
            If pos = 0 Then
                pos = InStr(myLine(0), "+")
                intSign = -1
            End If
            If myLine(0) = "Ostern" Then
                Holidays(intLine).Datum = Ostersonntag
            Else
            Select Case Left(myLine(0), pos - 1)
                Case "Ostern"
                   Holidays(intLine).Datum = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), Ostersonntag)
                Case "Advent"
                   Holidays(intLine).Datum = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), LetzterAdvent)
            End Select
            End If
            
        End If
        Holidays(intLine).Holiday = Trim(myLine(1))
        If InStr(myLine(2), ",") > 0 Then
            Dim c() As String
            c = Split(Trim(myLine(2)), ",")
            Dim a() As Variant
            ReDim a(UBound(c))
            Dim I As Integer
            For I = 0 To UBound(c)
                'a(i) = "{" & c(i) & "}"
                a(I) = c(I)
            Next
            Holidays(intLine).State = a
        Else
            Holidays(intLine).State = Array(Trim(myLine(2)))
        End If
        intLine = intLine + 1
    Loop
    Close #FNr
End Sub
Public Function isArrayEmpty(ByVal arr_var) As Boolean

    Dim p As Integer
    
    On Error Resume Next
        p = UBound(arr_var, 1)
    If Err.Number = 0 Then
        isArrayEmpty = False
    Else
        isArrayEmpty = True
    End If

End Function

Function IsInArray(stringToBeFound As Variant, arr As Variant) As Boolean
    IsInArray = Not IsError(Application.Match(stringToBeFound, arr, 0))
End Function

Private Function GetHoldidayByDate(ByVal Datum As Date, Optional ByVal State As String = vbNullString) As Holiday
    Dim h As Holiday
    Dim intCount As Integer
    For intCount = 0 To UBound(Holidays)
        If Holidays(intCount).Datum = Datum Then
            If State = "State" Then
                h = Holidays(intCount)
                Exit For
            End If
            If IsInArray("All", Holidays(intCount).State) = True Then
                h = Holidays(intCount)
                Exit For
            End If
            If IsInArray(State, Holidays(intCount).State) = True Then
                h = Holidays(intCount)
            End If
            Exit For
        End If
    Next
    GetHoldidayByDate = h
End Function

Public Function GetHolidayName(ByVal Datum As Date, Optional ByVal State As String = vbNullString) As String
    Dim h As Holiday
    If Year(Datum) <> Jahr Then
        Jahr = Year(Datum)
        InitDatumFile
    End If
    h = GetHoldidayByDate(Datum, State)
    GetHolidayName = h.Holiday
End Function

Public Function isHoliday(ByVal Datum As Date, _
    Optional ByVal CountryA As String = "de", _
    Optional ByVal State As String = vbNullString) As Boolean
    
    Dim h As Holiday
    If Year(Datum) <> Jahr Or CountryA <> Country Then
        Jahr = Year(Datum)
        Country = CountryA
        InitDatumFile
    End If
    h = GetHoldidayByDate(Datum, State)
    If h.Holiday = vbNullString Then
        isHoliday = False
    Else
        isHoliday = True
    End If
End Function

Public Function GetHolidayStates(ByVal Datum As Date) As String
    Dim h As Holiday
    If Year(Datum) <> Jahr Then
        Jahr = Year(Datum)
        InitDatumFile
    End If
    h = GetHoldidayByDate(Datum, "State")
    If h.Holiday = vbNullString Then
        GetHolidayStates = ""
    Else
        GetHolidayStates = Join(h.State, ", ")
        GetHolidayStates = Replace(Replace(GetHolidayStates, "}", ""), "{", "")
    End If
End Function

Public Function GetAllHolidays(ByVal JahrA As Integer, Optional ByVal CountryA As String = "de") As Variant()
    If JahrA <> Jahr Or CountryA <> Country Then
        Jahr = JahrA
        Country = CountryA
        InitDatumFile
    End If

    Dim intCount As Integer
    Dim HArray()
    For intCount = 0 To UBound(Holidays)
        ReDim Preserve HArray(intCount)
         
        HArray(intCount) = Holidays(intCount).Datum & "; " & Holidays(intCount).Holiday & "; " & Replace(Replace(Join(Holidays(intCount).State, ", "), "}", ""), "{", "")
    Next
    GetAllHolidays = HArray
End Function

Public Function GetCountries() As Variant()
    Dim intCount As Integer
    Dim CArray()
    For intCount = 0 To UBound(Countries)
        ReDim Preserve CArray(intCount)
         
        CArray(intCount) = Countries(intCount)
    Next
    GetCountries = CArray
End Function

Private Sub FillCountries()
    Dim strPath As String
    Dim strFile As String
    Dim intCount As Integer
    strPath = AddIns(strVBProjects).Path & "\countrycodes\"
    strFile = Dir(strPath & "*")
    Do While strFile <> ""
        ReDim Preserve Countries(intCount)
        Dim c() As String
        c = Split(strFile, ".")
        Countries(intCount) = c(0)
        intCount = intCount + 1
        strFile = Dir()
    Loop
End Sub
