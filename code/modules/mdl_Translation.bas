Attribute VB_Name = "mdl_Translation"
Option Explicit
Option Private Module

Public strRegister(12) As String
Public strCountry(3) As String

Public strErrorCaption As String
Public strErrorCaptionHint As String

Public Sub setTranslationStrings()
    strErrorCaption = t("Error")
    strErrorCaptionHint = t("Instering hint")
    'Function Easter
    strRegister(0) = t("GivenYear - Year for the Easter Sunday")
    strRegister(1) = t("Returns the date of Easter Sunday of the given year.")
    'Function LastAdvent
    strRegister(2) = t("GivenYear - Year for the 4th Advent Sunday")
    strRegister(3) = t("Returns the date of 4th Advent Sunday of the given year.")
    'Function isHoliday
    strRegister(4) = t("GivenDate - Date to be checked")
    strRegister(5) = t("Country - Country in 2-letter-ISO-Code for which the holiday shall be checked." _
        & " Default value is 'de'.")
    strRegister(6) = t("State - State for which the holiday shall be checked. (see documantation)" _
        & " No default value given.")
    strRegister(7) = t("Checks whether the given date is a holiday for a given country and tentative state.")
    'Function DayOfMonth
    strRegister(8) = t("GivenYear - year")
    strRegister(9) = t("GivenMonth - month given as number")
    strRegister(10) = t("DayOfWeek - given as number, 1 - Sunday to 7 - Saturday")
    strRegister(11) = t("NumInMonth - given as number, 1 - 5, 6 = last of month")
    strRegister(12) = t("Returns a date given by year, month, weekday and occurance in a month.")
    
    strCountry(0) = t("Germany")
    strCountry(1) = t("Austria")
    strCountry(2) = t("Switzerland")
    strCountry(3) = t("Italy")
End Sub



