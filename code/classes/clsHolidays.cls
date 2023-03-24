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

Private Type Holiday
    GivenDate As Date
    Holiday As String
    State() As Variant
End Type

Private Holidays() As Holiday

Private Sub Class_Initialize()
    GivenYear = Year(Now)
    Country = "de"
    InitDateFile
    FillCountries
End Sub

Private Sub InitDate()
    ReDim Holidays(16)
    Holidays(0).GivenDate = DateSerial(GivenYear, 1, 1)
    Holidays(0).Holiday = "NeuGivenYear"
    Holidays(0).State = Array("All")
    Holidays(1).GivenDate = DateSerial(GivenYear, 1, 6)
    Holidays(1).Holiday = "Drei Hl. Könige"
    Holidays(1).State = Array("BY", "ST", "BW")
    Holidays(2).GivenDate = DateSerial(GivenYear, 3, 8)
    Holidays(2).Holiday = "Int. Frauentag"
    Holidays(2).State = Array("BE")
    Dim Ostersonntag As Date
    Ostersonntag = Easter(GivenYear)
    Holidays(3).GivenDate = DateAdd("d", -2, Ostersonntag)
    Holidays(3).Holiday = "Karfreitag"
    Holidays(3).State = Array("All")
    Holidays(4).GivenDate = Ostersonntag
    Holidays(4).Holiday = "Ostersonntag"
    Holidays(4).State = Array("All")
    Holidays(5).GivenDate = DateAdd("d", 1, Ostersonntag)
    Holidays(5).Holiday = "Ostermontag"
    Holidays(5).State = Array("All")
    Holidays(6).GivenDate = DateAdd("d", 39, Ostersonntag)
    Holidays(6).Holiday = "Christi Himmelfahrt"
    Holidays(6).State = Array("All")
    Holidays(7).GivenDate = DateAdd("d", 50, Ostersonntag)
    Holidays(7).Holiday = "Pfingstmontag"
    Holidays(7).State = Array("All")
    Holidays(8).GivenDate = DateAdd("d", 60, Ostersonntag)
    Holidays(8).Holiday = "Fronleichnam"
    Holidays(8).State = Array("BW", "BY", "HE", "NW", "RP", "SL")
    Holidays(9).GivenDate = DateSerial(GivenYear, 8, 15)
    Holidays(9).Holiday = "Maria Himmelfahrt"
    Holidays(9).State = Array("BY", "SL")
    Holidays(10).GivenDate = DateSerial(GivenYear, 9, 20)
    Holidays(10).Holiday = "Weltkindertag"
    Holidays(10).State = Array("TH")
    Holidays(11).GivenDate = DateSerial(GivenYear, 10, 3)
    Holidays(11).Holiday = "Tag der dt. Einheit"
    Holidays(11).State = Array("All")
    Holidays(12).GivenDate = DateSerial(GivenYear, 10, 31)
    Holidays(12).Holiday = "Reformationstag"
    Holidays(12).State = Array("SH", "NI", "HB", "HH", "BB", "ST", "SN", "TH", "MV")
    Holidays(13).GivenDate = DateSerial(GivenYear, 11, 1)
    Holidays(13).Holiday = "Allerheiligen"
    Holidays(13).State = Array("BW", "BY", "NW", "RP", "SL")
    Dim LetzterAdvent As Date
    LetzterAdvent = LastAdvent(GivenYear)
    Holidays(14).GivenDate = DateAdd("d", -32, LetzterAdvent)
    Holidays(14).Holiday = "Buß- und Bettag"
    Holidays(14).State = Array("SN")
    Holidays(15).GivenDate = DateSerial(GivenYear, 12, 25)
    Holidays(15).Holiday = "1. Weihnachtstag"
    Holidays(15).State = Array("All")
    Holidays(16).GivenDate = DateSerial(GivenYear, 12, 26)
    Holidays(16).Holiday = "2. Weihnachtstag"
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
    Dim LastHoliday As Date
    Dim islYear As Long
    Dim tempHolidays() As Holiday
    Dim tempC As Integer
    islYear = getIslamicYear(DateSerial(GivenYear, 1, 1))
    Dim blnTemp As Boolean
    
    Dim myLine() As String
    FNr = FreeFile
    Dim strFile As String
    strFile = AddIns(strVBProjects).path & "\countrycodes\" & Country & ".inocsv"
    
    If mdlHelper_FileExists(strFile) = False Then
        InitDate
        Exit Sub
    End If
    
    LastHoliday = #1/1/2000#
    
    Open strFile For Input As FNr
    Do Until EOF(FNr)
        ReDim Preserve Holidays(intLine)
        Line Input #FNr, DataLine
        myLine = Split(DataLine, ";")
        
        Holidays(intLine).Holiday = VBA.Trim(myLine(1))
        If InStr(myLine(2), ",") > 0 Then
            Dim c() As String
            c = Split(VBA.Trim(myLine(2)), ",")
            Dim a() As Variant
            ReDim a(UBound(c))
            Dim i As Integer
            For i = 0 To UBound(c)
                'a(i) = "{" & c(i) & "}"
                a(i) = c(i)
            Next
            Holidays(intLine).State = a
        Else
            Holidays(intLine).State = Array(VBA.Trim(myLine(2)))
        End If
        
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
            If pos = 0 Then
                pos = InStr(myLine(0), ",")
            End If
            If myLine(0) = "Ostern" Then
                Holidays(intLine).GivenDate = Ostersonntag
            ElseIf Left(myLine(0), 3) = "Isl" Then
                Dim Isl() As String
                Dim islDate As Date
                Dim blnSecondEntry As Boolean
                
                Isl = Split(myLine(0), ",")
                islDate = getChristianDate(Isl(1) & " " & islYear & " AH")
                If islDate >= DateSerial(GivenYear, 1, 1) And islDate < DateSerial(GivenYear + 1, 1, 1) Then
                    Holidays(intLine).GivenDate = islDate
                    blnSecondEntry = True
                End If
                islDate = getChristianDate(Isl(1) & " " & islYear + 1 & " AH")
                If islDate >= DateSerial(GivenYear, 1, 1) And islDate < DateSerial(GivenYear + 1, 1, 1) Then
                    If blnSecondEntry = True Then
                        ReDim Preserve tempHolidays(tempC)
                        tempHolidays(tempC).GivenDate = islDate
                        tempHolidays(tempC).Holiday = Holidays(intLine).Holiday
                        tempHolidays(tempC).State = Holidays(intLine).State
                        tempC = tempC + 1
                        blnTemp = True
                    Else
                        Holidays(intLine).GivenDate = islDate
                    End If
                End If
            Else
                Select Case VBA.Left(myLine(0), pos - 1)
                    Case "Ostern"
                        Holidays(intLine).GivenDate = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), Ostersonntag)
                    Case "Advent"
                        Holidays(intLine).GivenDate = DateAdd("d", -intSign * CInt(Mid(myLine(0), pos + 1)), LetzterAdvent)
                    Case "Tag"
                        Dim Tage() As String
                        Tage = Split(myLine(0), ",")
                        Holidays(intLine).GivenDate = DayOfMonth(GivenYear, CInt(Tage(1)), CInt(Tage(2)), CInt(Tage(3)))
                End Select
            End If
            
        End If
        
        'Move holiday to monday if holiday on weekend
        If UBound(myLine) = 3 Then
            If VBA.LCase(VBA.Trim(myLine(3))) = "mo" Then
                If Weekday(Holidays(intLine).GivenDate) = vbSaturday Or Weekday(Holidays(intLine).GivenDate) = vbSunday Then
                    If LastHoliday = #1/1/2000# Then
                        If Weekday(Holidays(intLine).GivenDate) = vbSaturday Then
                            Holidays(intLine).GivenDate = DateAdd("d", 2, Holidays(intLine).GivenDate)
                        Else
                            Holidays(intLine).GivenDate = DateAdd("d", 1, Holidays(intLine).GivenDate)
                        End If
                    Else
                        If Weekday(Holidays(intLine).GivenDate) = vbSunday Then
                            Holidays(intLine).GivenDate = DateAdd("d", 2, Holidays(intLine).GivenDate)
                        End If
                    End If
                    LastHoliday = Holidays(intLine).GivenDate
                ElseIf LastHoliday > #1/1/2000# Then
                    If Weekday(Holidays(intLine - 1).GivenDate) = vbMonday And Holidays(intLine).GivenDate = Holidays(intLine - 1).GivenDate Then
                        Holidays(intLine).GivenDate = DateAdd("d", 1, Holidays(intLine).GivenDate)
                    End If
                Else
                    LastHoliday = #1/1/2000#
                End If
            End If
        End If
        

        intLine = intLine + 1
    Loop
    
    Close #FNr
    
    If blnTemp = True Then
        Dim intC As Integer
        For intC = 0 To tempC - 1
            ReDim Preserve Holidays(intLine)
            Holidays(intLine) = tempHolidays(intC)
            intLine = intLine + 1
        Next
    End If
    
    SortHolidays

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

Private Function GetHoldidayByDate(ByVal GivenDate As Date, Optional ByVal State As String = vbNullString) As Holiday
    Dim h As Holiday
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
    Dim h As Holiday
    If Year(GivenDate) <> GivenYear Then
        GivenYear = Year(GivenDate)
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, State)
    GetHolidayName = h.Holiday
End Function

Public Function isHoliday(ByVal GivenDate As Date, _
    Optional ByVal CountryA As String = "de", _
    Optional ByVal State As String = vbNullString) As Boolean
    
    Dim h As Holiday
    If Year(GivenDate) <> GivenYear Or CountryA <> Country Then
        GivenYear = Year(GivenDate)
        Country = CountryA
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, State)
    If h.Holiday = vbNullString Then
        isHoliday = False
    Else
        isHoliday = True
    End If
End Function

Public Function GetHolidayStates(ByVal GivenDate As Date) As String
    Dim h As Holiday
    If Year(GivenDate) <> GivenYear Then
        GivenYear = Year(GivenDate)
        InitDateFile
    End If
    h = GetHoldidayByDate(GivenDate, "State")
    If h.Holiday = vbNullString Then
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
         
        HArray(intCount) = Holidays(intCount).GivenDate & "; " & Holidays(intCount).Holiday & "; " & Replace(Replace(Join(Holidays(intCount).State, ", "), "}", ""), "{", "")
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
    strPath = AddIns(strVBProjects).path & "\countrycodes\"
    strFile = dir(strPath & "*")
    Do While strFile <> ""
        ReDim Preserve Countries(intCount)
        Dim c() As String
        c = Split(strFile, ".")
        Countries(intCount) = c(0)
        intCount = intCount + 1
        strFile = dir()
    Loop
End Sub

Private Sub SortHolidays()
    Dim TempH As Holiday
    Dim i As Integer, j As Integer
    For i = LBound(Holidays) To UBound(Holidays)
        For j = i + 1 To UBound(Holidays)
            If Holidays(i).GivenDate > Holidays(j).GivenDate Then
                TempH = Holidays(j)
                Holidays(j) = Holidays(i)
                Holidays(i) = TempH
            End If
        Next j
    Next i
End Sub
