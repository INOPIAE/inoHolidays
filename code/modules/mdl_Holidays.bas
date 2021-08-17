Attribute VB_Name = "mdl_Holidays"
Option Explicit

Private clsH As New clsHolidays

Public ImportRange As Range
Public ImportJahr As Integer
Public ImportBln As Boolean

Public Function Ostern(ByVal Jahr As Integer) As Date

    'ermittelt das Datum des Ostersonntags des ausgewählten Jahres
    
    Dim a, b, c, d, e, F, g, h, i, k, l, m, Wert, Monat, Datum
    
    a = Jahr Mod 19
    
    b = Jahr \ 100
    
    c = Jahr Mod 100
    
    d = b \ 4
    
    e = b Mod 4
    
    F = (b + 8) \ 25
    
    g = (b - F + 1) \ 3
    
    h = (19 * a + b - d - g + 15) Mod 30
    
    i = c \ 4
    
    k = c Mod 4
    
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    
    m = (a + 11 * h + 22 * l) \ 451
    
    Wert = h + l - 7 * m + 22
    
    Monat = 3 - (Wert > 31)
    
    Datum = Wert + 31 * (Wert > 31)
    
    Ostern = DateSerial(Jahr, Monat, Datum)

End Function

Public Function LetzterAdventSonntag(ByVal Jahr As Integer) As Date
    Dim dt As Date
    dt = DateSerial(Jahr, 12, 24)
    Dim wkday As Integer
    wkday = Weekday(dt, vbMonday)
    If wkday <> 7 Then
        dt = DateAdd("d", -wkday, dt)
    End If
    LetzterAdventSonntag = dt
End Function

Public Sub ImportHolidays()
    frmImportHoliday.Show
    
    If ImportBln = False Then Exit Sub
    
    Dim rng As Range
    Set rng = ImportRange
    
    Dim arr
    arr = clsH.GetAllHolidays(ImportJahr)

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




