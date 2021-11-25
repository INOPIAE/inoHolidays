Attribute VB_Name = "mdl_Translation"
Option Explicit
Option Private Module

Public strLabel(5) As String
Public strScreentip(5) As String
Public strSupertip(5) As String
Public strError(5) As String
Public strCmd(7) As String
Public strFrmInfo(1) As String
Public strFrmHolidays(7) As String
Public strFrmFunction(8) As String
Public strRegister(7) As String
Public strFrmOutlook(12) As String
Public strCountry(3) As String

Public Const strErrorCaption = "Error" ' < translate
Public Const strErrorCaptionHint = "Instering hint" ' < translate

Public Sub setTranslationStrings()
    strRegister(0) = t("GivenYear - Year for the Easter Sunday")
    strRegister(1) = t("Returns the date of Easter Sunday of the given year")
    strRegister(2) = t("GivenYear - Year for the 4th Advent Sunday")
    strRegister(3) = t("Returns the date of 4th Advent Sunday of the given year")
    strRegister(4) = t("GivenDate - Date to be checked")
    strRegister(5) = t("Country - Country in 2-letter-ISO-Code for which the holiday shall be checked." _
        & " Default value is 'de'.")
    strRegister(6) = t("State - State for which the holiday shall be checked. (see documantation)" _
        & " No default value given.")
    strRegister(7) = t("Checks whether the given date is a holiday for a given country and tentative state.")

    strCountry(0) = t("Germany")
    strCountry(1) = t("Austria")
    strCountry(2) = t("Switzerland")
    strCountry(3) = t("Italy")
End Sub



