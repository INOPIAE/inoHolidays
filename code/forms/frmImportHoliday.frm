VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportHoliday 
   Caption         =   "Feiertage importieren"
   ClientHeight    =   3555
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmImportHoliday.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmImportHoliday"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    ImportBln = False
    Unload Me
End Sub

Private Sub cmdImport_Click()
    If IsNumeric(Me.txtJahr.Value) = False Then
        MsgBox "Das Jahr muss als Zahl angegeben sein."
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    
    Dim rng As Range
    If Me.reZelle.Value = vbNullString Then
        MsgBox "Es muss eine Zelle ausgewählt werden."
        With Me.reZelle
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    Set rng = Range(Me.reZelle.Value)
    If rng.Cells.Count > 1 Then
        MsgBox "Es darf nur eine Zelle ausgewählt werden."
        With Me.reZelle
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    
    ImportBln = True
    ImportJahr = Me.txtJahr.Value
    Set ImportRange = rng
    Unload Me
End Sub
