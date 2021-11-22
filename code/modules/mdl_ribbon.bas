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
    SetLanguage
    Select Case control.ID
        Case "grpInoHolidays"
            label = "inoHolidays"
        Case "btnInoHolidays"
            label = strLabel(0)
        Case "btnInoOstern"
            label = strLabel(1)
        Case "btnInoLastAdvent"
            label = strLabel(2)
        Case "mnuInoRound"
            label = strLabel(3)
        Case "btnInfoInoHolidays"
            label = strLabel(4)
        Case "btnOutlookInoHolidays"
            label = strLabel(5)
        Case Else
            label = ""
    End Select
End Sub

Public Sub rbGetScreentipH(ByRef control As IRibbonControl, ByRef text)
    Select Case control.ID
        Case "btnInoHolidays"
            text = strScreentip(0)
        Case "btnInoOstern"
            text = strScreentip(1)
        Case "btnInoLastAdvent"
            text = strScreentip(2)
        Case "mnuInoRound"
            text = strScreentip(3)
        Case "btnInfoInoHolidays"
            text = strScreentip(4)
        Case "btnOutlookInoHolidays"
            text = strScreentip(5)
        Case Else
            text = ""
    End Select
End Sub

Public Sub rbGetSupertipH(ByRef control As IRibbonControl, ByRef text)
    Select Case control.ID
        Case "btnInoHolidays"
            text = strSupertip(0)
        Case "btnInoOstern"
            text = strSupertip(1)
        Case "btnInoLastAdvent"
            text = strSupertip(2)
        Case "mnuInoRound"
            text = strSupertip(3)
        Case "btnInfoInoHolidays"
            text = strSupertip(4)
        Case "btnOutlookInoHolidays"
            text = strSupertip(5)
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

Public Sub SetLanguage()
    lc = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    Select Case lc
        Case 1031
            germanText
        Case 1033
            englishText
        Case Else
            englishText
    End Select
End Sub
