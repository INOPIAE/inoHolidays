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

Public GivenYear As Integer
Public Country As String
Private Countries() As String

Private Type holiday
    GivenDate As Date
    holiday As String
    State() As Variant
End Type

Private Holidays() As holiday

Private Sub Class_Initialize()
    GivenYear = Year(Now)
    Country = "de"
    InitDateFile
    FillCountries
End Sub

Private Sub InitDate()
    ReDim Holidays(16)
    Holidays(0).GivenDate = DateSerial(GivenYear, 1, 1)
    Holidays(0).holiday = "NeuGivenYear"
    Holidays(0).State = Array("All")
    Holidays(1).GivenDate = DateSerial(GivenYear, 1, 6)
    Holidays(1).holiday = "Drei Hl. Könige"
    Holidays(1).State = Array("BY", "ST", "BW")
    Holidays(2).GivenDate = DateSerial(GivenYear, 3, 8)
    Holidays(2).holiday = "Int. Frauentag"
    Holidays(2).State = Array("BE")
    Dim Ostersonntag As Date
    Ostersonntag = Easter(GivenYear)
    Holidays(3).GivenDate = DateAdd("d", -2, Ostersonntag)
    Holidays(3).holiday = "Karfreitag"
    Holidays(3).State = Array("All")
    Holidays(4).GivenDate = Ostersonntag
    Holidays(4).holiday = "Ostersonntag"
    Holidays(4).State = Array("All")
    Holidays(5).GivenDate = DateAdd("d", 1, Ostersonntag)
    Holidays(5).holiday = "Ostermontag"
    Holidays(5).State = Array("All")
    Holidays(6).GivenDate = DateAdd("d", 39, Ostersonntag)
    Holidays(6).holiday = "Christi Himmelfahrt"
    Holidays(6).State = Array("All")
    Holidays(7).GivenDate = DateAdd("d", 50, Ostersonntag)
    Holidays(7).holiday = "Pfingstmontag"
    Holidays(7).State = Array("All")
    Holidays(8).GivenDate = DateAdd("d", 60, Ostersonntag)
    Holidays(8).holiday = "Fronleichnam"
    Holidays(8).State = Array("BW", "BY", "HE", "NW", "RP", "SL")
    Holidays(9).GivenDate = DateSerial(GivenYear, 8, 15)
    Holidays(9).holiday = "Maria Himmelfahrt"
    Holidays(9).State = Array("BY", "SL")
    Holidays(10).GivenDate = DateSerial(GivenYear, 9, 20)
    Holidays(10).holiday = "Weltkindertag"
    Holidays(10).State = Array("TH")
    Holidays(11).GivenDate = DateSerial(GivenYear, 10, 3)
    Holidays(11).holiday = "Tag der dt. Einheit"
    Holidays(11).State = Array("All")
    Holidays(12).GivenDate = DateSerial(GivenYear, 10, 31)
    Holidays(12).holiday = "Reformationstag"
    Holidays(12).State = Array("SH", "NI", "HB", "HH", "BB", "ST", "SN", "TH", "MV")
    Holidays(13).GivenDate = DateSerial(GivenYear, 11, 1)
    Holidays(13).holiday = "Allerheiligen"
    Holidays(13).State = Array("BW", "BY", "NW", "RP", "SL")
    Dim LetzterAdvent As Date
    LetzterAdvent = LastAdvent(GivenYear)
    Holidays(14).GivenDate = DateAdd("d", -32, LetzterAdvent)
    Holidays(14).holiday = "Buß- und Bettag"
    Holidays(14).State = Array("SN")
    Holidays(15).GivenDate = DateSerial(GivenYear, 12, 25)
    Holidays(15).holiday = "1. Weihnachtstag"
    Holidays(15).State = Array("All")
    Holidays(16).GivenDate = DateSerial(GivenYear, 12, 26)
    Holidays(16).holiday = "2. Weihnachtstag"
    Holidays(16).State = Array("All")

End Sub

Private Sub InitDateFile()
    Dim FNr As Integer
    Dim intLine As Integer
    Dim DataLine As String
    Dim Ostersonntag As Date
    Ostersonntag = Easter(GivenYear)
    Dim LetzterAdvent As Date
    LetzterAdvent = LastAdvent(GivenYear)

    Dim myLine() As String
    FNr = FreeFile
    Dim strFile As String
    strFile = AddIns(strVBProjects).Path & "\countrycodes\" & Country & ".inocsv"
    
    If mdlHelper_FileExists(strFile) = False Then
        InitDate
        Exit Sub
    End If
    
    Open strFile For Input As FNr
    Do Until EOF(FNr)
        ReDim Preserve Holidays(intLine)
        Line Input #FNr, DataLine
        myLine = Split(DataLine, ";")
        If IsDate(myLine(0)) Then
            Holidays(intLine).GivenDate = DateSerial(GivenYear, Month(myLine(0)), Day(myLine(0)))
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
                Holidays(intLine).GivenDate = Ostersonntag
            Else
            Select Case Left(myLine(0), pos - 1)
                Case "Ostern"
                   Holidays(intLine).GivenDate = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), Ostersonntag)
                Case "Advent"
                   Holidays(intLine).GivenDate = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), LetzterAdvent)
            End Select
            End If
            
        End If
        Holidays(intLine).holiday = Trim(myLine(1))
        If InStr(myLine(2), ",") > 0 Then
            Dim c() As String
            c = Split(Trim(myLine(2)), ",")
            Dim a() As Variant
            ReDim a(UBound(c))
            Dim i As Integer
            For i = 0 To UBound(c)
                'a(i) = "{" & c(i) & "}"
                a(i) = c(i)
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

Private Function GetHoldidayByDate(ByVal GivenDate As Date, Optional ByVal State As String = vbNullString) As holiday
    Dim h As holiday
    Dim intCount As Integer
    For intCount = 0 To UBound(Holidays)
        If Holidays(intCount).GivenDate = GivenDate Then
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

Public Function GetHolidayName(ByVal GivenDate As Date, Optional ByVal State As String = vbNullString) As String
    Dim h As holiday
    If Year(GivenDate) <> GivenYear Then
        GivenYear = Year(GivenDate)
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, State)
    GetHolidayName = h.holiday
End Function

Public Function isHoliday(ByVal GivenDate As Date, _
    Optional ByVal CountryA As String = "de", _
    Optional ByVal State As String = vbNullString) As Boolean
    
    Dim h As holiday
    If Year(GivenDate) <> GivenYear Or CountryA <> Country Then
        GivenYear = Year(GivenDate)
        Country = CountryA
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, State)
    If h.holiday = vbNullString Then
        isHoliday = False
    Else
        isHoliday = True
    End If
End Function

Public Function GetHolidayStates(ByVal GivenDate As Date) As String
    Dim h As holiday
    If Year(GivenDate) <> GivenYear Then
        GivenYear = Year(GivenDate)
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, "State")
    If h.holiday = vbNullString Then
        GetHolidayStates = ""
    Else
        GetHolidayStates = Join(h.State, ", ")
        GetHolidayStates = Replace(Replace(GetHolidayStates, "}", ""), "{", "")
    End If
End Function

Public Function GetAllHolidays(ByVal GivenYearA As Integer, Optional ByVal CountryA As String = "de") As Variant()
    If GivenYearA <> GivenYear Or CountryA <> Country Then
        GivenYear = GivenYearA
        Country = CountryA
        InitDateFile
    End If

    Dim intCount As Integer
    Dim HArray()
    For intCount = 0 To UBound(Holidays)
        ReDim Preserve HArray(intCount)
         
        HArray(intCount) = Holidays(intCount).GivenDate & "; " & Holidays(intCount).holiday & "; " & Replace(Replace(Join(Holidays(intCount).State, ", "), "}", ""), "{", "")
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
