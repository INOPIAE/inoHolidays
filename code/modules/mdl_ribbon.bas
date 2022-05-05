Attribute VB_Name = "mdl_ribbon"
Option Explicit
Option Private Module

Public lc As Long
Public gobjRibbon As IRibbonUI

'Ribbon functions
Public Sub OnRibbonLoad(ByRef ribbon As IRibbonUI)
    Set gobjRibbon = ribbon
End Sub

Public Sub rbGetLabelH(ByRef control As IRibbonControl, ByRef label As Variant)
'    rbSetLanguage
    Select Case control.id
        Case "grpInoHolidays"
            label = "inoHolidays"
        Case "btnInoHolidays"
            label = t("Import Holidays")
        Case "btnInoOstern"
            label = t("Function Easter")
        Case "btnInoLastAdvent"
            label = t("Function LastAdvent")
        Case "mnuInoRound"
            label = ""
        Case "btnInfoInoHolidays"
            label = t("Info")
        Case "btnOutlookInoHolidays"
            label = t("Add Holidays to Outlook")
        Case "btnProjectInoHolidays"
            label = t("Add Holidays to MS Project")
        Case Else
            label = ""
    End Select
End Sub

Public Sub rbGetScreentipH(ByRef control As IRibbonControl, ByRef text)
    Select Case control.id
        Case "btnInoHolidays"
            text = t("Import holidays of a given Year.")
        Case "btnInoOstern"
            text = t("Function Easter returns the date of Easter sunday of a given year.")
        Case "btnInoLastAdvent"
            text = t("Function LastAdvent returns the date of the 4th Sunday in Advent of a given year.")
        Case "mnuInoRound"
            text = ""
        Case "btnInfoInoHolidays"
            text = ""
        Case "btnOutlookInoHolidays"
            text = t("Add Holidays to Outlook")
        Case "btnProjectInoHolidays"
            text = t("Add Holidays to MS Project")
        Case Else
            text = ""
    End Select
End Sub

Public Sub rbGetSupertipH(ByRef control As IRibbonControl, ByRef text)
    Select Case control.id
        Case "btnInoHolidays"
            text = t("Import holidays of a given Year.")
        Case "btnInoOstern"
            text = t("Function Easter returns the date of Easter sunday of a given year.")
        Case "btnInoLastAdvent"
            text = t("Function LastAdvent returns the date of the 4th Sunday in Advent of a given year.")
        Case "mnuInoRound"
            text = ""
        Case "btnInfoInoHolidays"
            text = ""
        Case "btnOutlookInoHolidays"
            text = t("Add Holidays to Outlook")
        Case "btnProjectInoHolidays"
            text = t("Add Holidays to MS Project")
        Case Else
            text = ""
    End Select
End Sub

' control functions
Public Sub rbOstern(ctrl As IRibbonControl)
    With frmFunction
        .InitForm ("Ostern")
        .Show
    End With
End Sub

Public Sub rbAdvent(ctrl As IRibbonControl)
    With frmFunction
        .InitForm ("Advent")
        .Show
    End With
End Sub

Public Sub rbHolidays(ctrl As IRibbonControl)
    ImportHolidays
End Sub

Public Sub rbInfoInoHolidays(ctrl As IRibbonControl)
    frm_Info.Show
End Sub

Public Sub rbOutlookInoHolidays(ctrl As IRibbonControl)
    If CheckVBAReferences("Outlook") = True Then
        frmOutlookImport.Show
    End If
End Sub


Public Sub rbProjectInoHolidays(ctrl As IRibbonControl)
    If CheckVBAReferences("MSProject") = True Then
        ShowProjectImport
    End If
End Sub


Public Sub rbSetLanguage()
'    lc = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
'    Select Case lc
'        Case 1031
'            germanText
'        Case 1033
'            englishText
'        Case Else
'            englishText
'    End Select
End Sub
