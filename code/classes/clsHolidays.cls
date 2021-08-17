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

Private Type Holiday
    Datum As Date
    Holiday As String
    State() As Variant
End Type

Private Holidays() As Holiday

Private Sub Class_Initialize()
    Jahr = Year(Now)
    InitDatum
End Sub

Private Sub InitDatum()
    ReDim Holidays(17)
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
    Holidays(7).Datum = DateAdd("d", 39, Ostersonntag)
    Holidays(7).Holiday = "Christi Himmelfahrt"
    Holidays(7).State = Array("All")
    Holidays(8).Datum = DateAdd("d", 50, Ostersonntag)
    Holidays(8).Holiday = "Pfingstmontag"
    Holidays(8).State = Array("All")
    Holidays(9).Datum = DateAdd("d", 60, Ostersonntag)
    Holidays(9).Holiday = "Fronleichnam"
    Holidays(9).State = Array("BW", "BY", "HE", "NW", "RP", "SL")
    Holidays(10).Datum = DateSerial(Jahr, 8, 15)
    Holidays(10).Holiday = "Maria Himmelfahrt"
    Holidays(10).State = Array("BY", "SL")
    Holidays(11).Datum = DateSerial(Jahr, 9, 20)
    Holidays(11).Holiday = "Weltkindertag"
    Holidays(11).State = Array("TH")
    Holidays(12).Datum = DateSerial(Jahr, 10, 3)
    Holidays(12).Holiday = "Tag der dt. Einheit"
    Holidays(12).State = Array("All")
    Holidays(13).Datum = DateSerial(Jahr, 10, 31)
    Holidays(13).Holiday = "Reformationstag"
    Holidays(13).State = Array("SH", "NI", "HB", "HH", "BB", "ST", "SN", "TH", "MV")
    Holidays(14).Datum = DateSerial(Jahr, 11, 1)
    Holidays(14).Holiday = "Allerheiligen"
    Holidays(14).State = Array("BW", "BY", "NW", "RP", "SL")
    Dim LetzterAdvent As Date
    LetzterAdvent = LetzterAdventSonntag(Jahr)
    Holidays(15).Datum = DateAdd("d", -32, LetzterAdvent)
    Holidays(15).Holiday = "Buß- und Bettag"
    Holidays(15).State = Array("SN")
    Holidays(16).Datum = DateSerial(Jahr, 12, 25)
    Holidays(16).Holiday = "1. Weihnachtstag"
    Holidays(16).State = Array("All")
    Holidays(17).Datum = DateSerial(Jahr, 12, 26)
    Holidays(17).Holiday = "2. Weihnachtstag"
    Holidays(17).State = Array("All")

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
        InitDatum
    End If
    h = GetHoldidayByDate(Datum, State)
    GetHolidayName = h.Holiday
End Function

Public Function IsHoliday(ByVal Datum As Date, Optional ByVal State As String = vbNullString) As Boolean
    Dim h As Holiday
    If Year(Datum) <> Jahr Then
        Jahr = Year(Datum)
        InitDatum
    End If
    h = GetHoldidayByDate(Datum, State)
    If h.Holiday = vbNullString Then
        IsHoliday = False
    Else
        IsHoliday = True
    End If
End Function

Public Function GetHolidayStates(ByVal Datum As Date) As String
    Dim h As Holiday
    If Year(Datum) <> Jahr Then
        Jahr = Year(Datum)
        InitDatum
    End If
    h = GetHoldidayByDate(Datum, "State")
    If h.Holiday = vbNullString Then
        GetHolidayStates = ""
    Else
        GetHolidayStates = Join(h.State, ", ")
    End If
End Function

Public Function GetAllHolidays(ByVal JahrA As Integer) As Variant()
    If JahrA <> Jahr Then
        Jahr = JahrA
        InitDatum
    End If

    Dim intCount As Integer
    Dim HArray()
    For intCount = 0 To UBound(Holidays)
        ReDim Preserve HArray(intCount)
         
        HArray(intCount) = Holidays(intCount).Datum & "; " & Holidays(intCount).Holiday & "; " & Join(Holidays(intCount).State, ", ")
    Next
    GetAllHolidays = HArray
End Function
