Attribute VB_Name = "mdl_ExposedFunctions"
Option Explicit

Public Function Easter(ByVal GivenYear As Integer) As Date
    Easter = IEaster(ByVal GivenYear)
End Function

Public Function LastAdvent(ByVal GivenYear As Integer) As Date
    LastAdvent = ILastAdvent(GivenYear)
End Function

Public Function isHoliday(ByVal GivenDate As Date, _
    Optional ByVal Country As String = "de", Optional ByVal State As String = vbNullString) As Boolean
    isHoliday = IisHoliday(GivenDate, Country, State)
End Function

Public Function DayOfMonth(ByVal GivenYear As Integer, GivenMonth As Integer, ByVal DayOfWeek As VbDayOfWeek, ByVal NumInMonth As NumberInMonth) As Variant
    DayOfMonth = IDayOfMonth(GivenYear, GivenMonth, DayOfWeek, NumInMonth)
End Function

Function getChristianDate(ByVal GivenIslamicDate As String) As Date
    getChristianDate = IgetChristianDate(GivenIslamicDate)
End Function

Function getIslamicDate(ByVal GivenDate As Date) As String
    getIslamicDate = IgetIslamicDate(GivenDate)
End Function
