Attribute VB_Name = "mdl_ExposedFunctions"
Option Explicit

Public Function Easter(ByVal GivenYear As Integer) As Date
Attribute Easter.VB_Description = "Gibt das Datum des Ostersonntags für das angegebene Jahr zurück."
Attribute Easter.VB_ProcData.VB_Invoke_Func = " \n20"
    Easter = IEaster(ByVal GivenYear)
End Function

Public Function LastAdvent(ByVal GivenYear As Integer) As Date
Attribute LastAdvent.VB_Description = "Gibt das Datum des 4. Adventsonntags für das angegebene Jahr zurück."
Attribute LastAdvent.VB_ProcData.VB_Invoke_Func = " \n20"
    LastAdvent = ILastAdvent(GivenYear)
End Function

Public Function isHoliday(ByVal GivenDate As Date, _
    Optional ByVal Country As String = "de", Optional ByVal State As String = vbNullString) As Boolean
Attribute isHoliday.VB_Description = "Prüft, ob das angegebene Datum (GivenDate) unter Berücksichtigung des Staates und evtl. Bundeslandes ein Feiertag ist."
Attribute isHoliday.VB_ProcData.VB_Invoke_Func = " \n20"
    isHoliday = IisHoliday(GivenDate, Country, State)
End Function

Public Function DayOfMonth(ByVal GivenYear As Integer, GivenMonth As Integer, ByVal DayOfWeek As VbDayOfWeek, ByVal NumInMonth As NumberInMonth) As Variant
Attribute DayOfMonth.VB_Description = "Gibt das Datum für die Eingabe von Jahr, Monat, Wochentag und Vorkommen im Monat zurück."
Attribute DayOfMonth.VB_ProcData.VB_Invoke_Func = " \n20"
    DayOfMonth = IDayOfMonth(GivenYear, GivenMonth, DayOfWeek, NumInMonth)
End Function

Function getChristianDate(ByVal GivenIslamicDate As String) As Date
Attribute getChristianDate.VB_Description = "Gibt ein Christliches Datum aus  einem islamischen zurück, z. B. '1. Muharram 1445 AH'.  Schreibweise der islamischen Monate:\r\nMuharram, Safar, Rabi I, Rabi II, Jumada I, Jumada II, Radschab, Sha'ban, Ramadan, Schawwal, Dhu'l-Qa'dah, Dhu'l-Hijja"
Attribute getChristianDate.VB_ProcData.VB_Invoke_Func = " \n20"
    getChristianDate = IgetChristianDate(GivenIslamicDate)
End Function

Function getIslamicDate(ByVal GivenDate As Date) As String
Attribute getIslamicDate.VB_Description = "Gibt ein islamisches Datum aus einem  christlichen Datum zurück."
Attribute getIslamicDate.VB_ProcData.VB_Invoke_Func = " \n20"
    getIslamicDate = IgetIslamicDate(GivenDate)
End Function
