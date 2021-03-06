Attribute VB_Name = "mdl_Holidays"
Option Explicit

Private clsH As New clsHolidays

Public ImportRange As Range
Public ImportGivenYear As Integer
Public ImportBln As Boolean
Public ImportCountry As String

Public Enum NumberInMonth
    First = 1
    Second = 2
    Third = 3
    Fourth = 4
    Fifth = 5
    Last = 6
End Enum

Public Function Easter(ByVal GivenYear As Integer) As Date
Attribute Easter.VB_Description = "Gibt das Datum des Ostersonntags f?r das angegebene Jahr zur?ck."
Attribute Easter.VB_ProcData.VB_Invoke_Func = " \n20"

    'calculates the date of Easter of a given year
    
    Dim a, b, c, d, e, f, g, h, i, k, l, m, W, Mon, GivenDate
    
    a = GivenYear Mod 19
    
    b = GivenYear \ 100
    
    c = GivenYear Mod 100
    
    d = b \ 4
    
    e = b Mod 4
    
    f = (b + 8) \ 25
    
    g = (b - f + 1) \ 3
    
    h = (19 * a + b - d - g + 15) Mod 30
    
    i = c \ 4
    
    k = c Mod 4
    
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    
    m = (a + 11 * h + 22 * l) \ 451
    
    W = h + l - 7 * m + 22
    
    Mon = 3 - (W > 31)
    
    GivenDate = W + 31 * (W > 31)
    
    Easter = DateSerial(GivenYear, Mon, GivenDate)

End Function

Public Function LastAdvent(ByVal GivenYear As Integer) As Date
Attribute LastAdvent.VB_Description = "Gibt das Datum des 4. Adventsonntags f?r das angegebene Jahr zur?ck."
Attribute LastAdvent.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim dt As Date
    dt = DateSerial(GivenYear, 12, 24)
    Dim wkday As Integer
    wkday = Weekday(dt, vbMonday)
    If wkday <> 7 Then
        dt = DateAdd("d", -wkday, dt)
    End If
    LastAdvent = dt
End Function

Public Sub ImportHolidays()
    With frmImportHoliday
        .FillForm clsH.GivenYear, clsH.Country
        .Show
    End With
    If ImportBln = False Then Exit Sub
    
    Dim rng As Range
    Set rng = ImportRange
    
    Dim arr
    arr = clsH.GetAllHolidays(ImportGivenYear, ImportCountry)

    rng.Resize(UBound(arr) + 1, 1).Value = Application.Transpose(arr) ' arr
    Application.DisplayAlerts = False
    Set rng = Range(rng, Cells(rng.Row + UBound(arr), rng.Column))
    
    rng.TextToColumns Destination:=rng, DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
            Semicolon:=True, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(Array(1, 4), Array(2, 1), Array(3, 1)), TrailingMinusNumbers:=True
    Application.DisplayAlerts = True
    
    Set rng = rng.Resize(rng.Rows.Count, 3)
    rng.Columns.AutoFit
    
End Sub

Public Function isHoliday(ByVal GivenDate As Date, _
    Optional ByVal Country As String = "de", Optional ByVal State As String = vbNullString) As Boolean
Attribute isHoliday.VB_Description = "Pr?ft, ob das angegebene Datum (GivenDate) unter Ber?cksichtigung des Staates und evtl. Bundeslandes ein Feiertag ist."
Attribute isHoliday.VB_ProcData.VB_Invoke_Func = " \n20"
    isHoliday = clsH.isHoliday(GivenDate, Country, State)
End Function

Public Function DayOfMonth(ByVal GivenYear As Integer, GivenMonth As Integer, ByVal DayOfWeek As VbDayOfWeek, ByVal NumInMonth As NumberInMonth) As Variant
Attribute DayOfMonth.VB_Description = "Gibt das Datum f?r die Eingabe von Jahr, Monat, Wochentag und Vorkommen im Monat zur?ck."
Attribute DayOfMonth.VB_ProcData.VB_Invoke_Func = " \n20"
    Dim PosMonth As Integer
    Dim GivenDate As Date
    GivenDate = DateSerial(GivenYear, GivenMonth, 1)
    PosMonth = NumInMonth
    If NumInMonth = 6 Then PosMonth = NumInMonth - 1
    
    DayOfMonth = GivenDate - Weekday(GivenDate - DayOfWeek + 1, 2) + 7 * PosMonth

    If Month(DayOfMonth) <> Month(GivenDate) Then
        If NumInMonth = 6 Then
            PosMonth = NumInMonth - 2
            DayOfMonth = GivenDate - Weekday(GivenDate - DayOfWeek + 1, 2) + 7 * PosMonth
        Else
            DayOfMonth = CVErr(xlErrNA)
        End If
    End If
End Function
