Attribute VB_Name = "mdl_IslamicCalender"
Option Explicit


' 0 - running days of years, 1 - is islamic leap year: 0 - no leap year, 1 -leap year
Private IslamicYears(30, 1) As Integer
' 0 - monthname, 1 - days in month, 2 - running total of days of year
Private IslamicMonth(12, 2)

Private Sub fillVariables()
    Dim intC As Integer
    Dim intDays As Integer

    If IslamicYears(1, 0) = 0 Then
        For intC = 1 To 30
            Select Case intC
                Case 2, 5, 7, 10, 13, 16, 18, 21, 24, 26, 29
                    intDays = intDays + 355
                    IslamicYears(intC, 1) = 1
                Case Else
                    intDays = intDays + 354
                    IslamicYears(intC, 1) = 0
            End Select
            IslamicYears(intC, 0) = intDays
        Next
        IslamicYears(0, 0) = 0
        intDays = 0
        For intC = 1 To 12
            intDays = 29 + intC Mod 2
            IslamicMonth(intC, 1) = intDays
            IslamicMonth(intC, 2) = intDays + IslamicMonth(intC - 1, 2)
        Next
        IslamicMonth(0, 1) = 0
        IslamicMonth(0, 2) = 0
        IslamicMonth(1, 0) = "Muharram"
        IslamicMonth(2, 0) = "Safar"
        IslamicMonth(3, 0) = "Rabi I"
        IslamicMonth(4, 0) = "Rabi II"
        IslamicMonth(5, 0) = "Jumada I"
        IslamicMonth(6, 0) = "Jumada II"
        IslamicMonth(7, 0) = "Radschab"
        IslamicMonth(8, 0) = "Sha'ban"
        IslamicMonth(9, 0) = "Ramadan"
        IslamicMonth(10, 0) = "Schawwal"
        IslamicMonth(11, 0) = "Dhu'l-Qa'dah"
        IslamicMonth(12, 0) = "Dhu'l-Hijja"
   End If

End Sub

Private Function getIslamicDay(ByVal GivenDate As Date) As Long
    Dim js As Long
    Dim jsj As Long
    Dim jsd As Long
    Dim jdate As Date
    
    jdate = DateAdd("d", -12, GivenDate)
    js = (Year(jdate) \ 4) * 1461
    jsj = ((Year(jdate) Mod 4) - 1) * 365
    jsd = DateDiff("d", DateSerial(Year(jdate), 1, 1), jdate)
    getIslamicDay = js + jsj + jsd - 227016 - IIf(Year(jdate) Mod 4 = 0, 1, 0)
End Function

Public Function getIslamicYear(ByVal GivenDate As Date) As Long
    Dim IDay As Long
    Dim isc As Long
    Dim isr As Long
    Dim intC As Long
    fillVariables
    IDay = getIslamicDay(GivenDate)
    
    isc = IDay \ 10631
    isr = IDay Mod 10631
    
    For intC = 0 To 30
        If IslamicYears(intC, 0) > isr Then
            Exit For
        End If
    Next
    If isr - IslamicYears(intC - 1, 0) = 0 Then
        getIslamicYear = isc * 30 + intC - 1
    Else
        getIslamicYear = isc * 30 + intC
    End If
End Function


Private Function getIslamicMonth(ByVal GivenDate As Date) As String
    Dim IDay As Long
    Dim isc As Long
    Dim isr As Long
    Dim intC As Long
    Dim ism As Long
    fillVariables
    IDay = getIslamicDay(GivenDate)
    
    isc = IDay \ 10631
    isr = IDay Mod 10631
    
    For intC = 1 To 30
        If IslamicYears(intC, 0) > isr Then
            Exit For
        End If
    Next
    ism = isr - IslamicYears(intC - 1, 0)
    For intC = 1 To 12
        If IslamicMonth(intC, 2) >= ism Then
            Exit For
        End If
    Next
    If intC = 13 Then intC = 12
    If ism = 0 Then
        getIslamicMonth = IslamicMonth(12, 0)
    Else
        getIslamicMonth = IslamicMonth(intC, 0)
    End If
End Function

Private Function getIslamicDayOfMonth(ByVal GivenDate As Date) As Long
    Dim IDay As Long
    Dim isc As Long
    Dim isr As Long
    Dim intC As Long
    Dim intC1 As Long
    Dim ism As Long
    Dim IslamicDayOfMonthK As Long
    
    fillVariables
    IDay = getIslamicDay(GivenDate)

StartCorrection:
    isc = IDay \ 10631
    isr = IDay Mod 10631
    
    For intC = 1 To 30
        If IslamicYears(intC, 0) > isr Then
            Exit For
        End If
    Next

    ism = isr - IslamicYears(intC - 1, 0)
    For intC1 = 1 To 12
        If IslamicMonth(intC1, 2) > ism Then
            Exit For
        End If
    Next
    Dim IslamicDayOfMonth
    If ism = 0 Then
        IslamicDayOfMonth = 29
    Else
        IslamicDayOfMonth = ism - IslamicMonth(intC1 - 1, 2)
    End If
    
    If IslamicDayOfMonth = 0 Then
        IslamicDayOfMonth = IslamicMonth(intC1 - 1, 1)
    End If

    If intC1 = 1 And IslamicYears(intC - 1, 1) = 1 Then
        
        If IslamicDayOfMonth = 29 And IslamicDayOfMonthK = 0 Then
            IslamicDayOfMonthK = IslamicDayOfMonth
            IDay = IDay - 1
            GoTo StartCorrection
        Else
            If IslamicDayOfMonthK > 0 Then
                If IslamicDayOfMonth = IslamicDayOfMonthK Then
                    IslamicDayOfMonth = 30
                Else
                    IslamicDayOfMonth = 29
                End If
            End If
        End If
    End If
    If intC1 = 13 And IslamicYears(intC - 1, 1) = 0 And IslamicDayOfMonthK > 0 Then
        If IslamicDayOfMonth = IslamicDayOfMonthK Then
            IslamicDayOfMonth = 30
        Else
            IslamicDayOfMonth = 29
        End If
    End If
    getIslamicDayOfMonth = IslamicDayOfMonth
End Function

Function getIslamicDate(ByVal GivenDate As Date) As String
Attribute getIslamicDate.VB_Description = "Gibt ein islamisches Datum aus einem  christlichen Datum zurück."
Attribute getIslamicDate.VB_ProcData.VB_Invoke_Func = " \n14"
    getIslamicDate = getIslamicDayOfMonth(GivenDate) & ". " & getIslamicMonth(GivenDate) _
        & " " & getIslamicYear(GivenDate) & " AH"
End Function

Private Function getJulianDays(ByVal GivenIslamicDate As String) As Long
    Dim strT() As String
    GivenIslamicDate = Replace(GivenIslamicDate, " I", "_I")
    strT = Split(Trim(GivenIslamicDate), " ")
    fillVariables
    Dim isy As Long
    isy = CLng(strT(2)) \ 30
    Dim isyr As Long
    isyr = (CLng(strT(2)) Mod 30)
    
    Dim intC As Integer
    
    For intC = 1 To 12
        If IslamicMonth(intC, 0) = Replace(Mid(strT(1), 1), "_I", " I") Then
            Exit For
        End If
    Next
    If isyr = 0 Then
        getJulianDays = isy * 10631 - 354 + IslamicMonth(intC - 1, 2) + CLng(strT(0)) - 1 + 227016
    Else
        getJulianDays = isy * 10631 + IslamicYears(isyr - 1, 0) + IslamicMonth(intC - 1, 2) + CLng(strT(0)) - 1 + 227016
    End If
End Function

Function getChristianDate(ByVal GivenIslamicDate As String) As Date
Attribute getChristianDate.VB_Description = "Gibt ein Christliches Datum aus  einem islamischen zurück, z. B. '1. Muharram 1445 AH'.  Schreibweise der islamischen Monate:\r\nMuharram, Safar, Rabi I, Rabi II, Jumada I, Jumada II, Radschab, Sha'ban, Ramadan, Schawwal, Dhu'l-Qa'dah, Dhu'l-Hijja"
Attribute getChristianDate.VB_ProcData.VB_Invoke_Func = " \n14"
    fillVariables
    Dim id As Long
    Dim chrDateK As Date
    chrDateK = #1/1/1900#
    
    id = getJulianDays(GivenIslamicDate)
StartCorrection:
    Dim ics As Long
    Dim icsr As Long
    Dim icy As Long
    Dim icyr As Long
    
    ics = id \ 1461
    icsr = id Mod 1461
    icy = icsr \ 365
    icyr = icsr Mod 365
    
    Dim bDate As Date
    Dim chrDate As Date
    bDate = DateSerial(ics * 4 + icy + 1, 1, 1)

    chrDate = DateAdd("d", icyr + 13, bDate)

    If Month(chrDate) = 1 And Day(chrDate) = 14 Then
        If chrDateK = #1/1/1900# Then
            chrDateK = chrDate
            id = id + 1
            GoTo StartCorrection
        Else
            If chrDate = chrDateK Then
                chrDate = DateAdd("d", -1, chrDate)
            End If
        End If
    Else
        If chrDateK <> #1/1/1900# Then
            chrDate = chrDateK
        End If
    End If
    getChristianDate = chrDate
End Function
