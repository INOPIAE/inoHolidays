Attribute VB_Name = "mdl_ribbon"
Option Explicit

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

Public Sub rbFeiertage(ctrl As IRibbonControl)
    ImportHolidays
End Sub

Public Sub rbInfoInoHolidays(ctrl As IRibbonControl)
    frm_Info.Show
End Sub
