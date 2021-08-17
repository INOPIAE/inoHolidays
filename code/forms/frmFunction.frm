VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFunction 
   Caption         =   "UserForm1"
   ClientHeight    =   4050
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7380
   OleObjectBlob   =   "frmFunction.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rngStart As Range
Private rngValue As Range
Public strFunction As String
Private intJahr As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    If intJahr > 0 Then
        Select Case strFunction
            Case "Ostern"
                rngStart.Formula = "=Ostern(" & Me.reJahr.Value & ")"
            Case "Advent"
                rngStart.Formula = "=LetzterAdventSonntag(" & Me.reJahr.Value & ")"
        End Select
        rngStart.NumberFormat = "dd.MM.yyyy"
        Unload Me
    End If
End Sub

Private Sub cmdValue_Click()
    If intJahr > 0 Then
        Select Case strFunction
            Case "Ostern"
                rngStart.Value = Ostern(intJahr)
            Case "Advent"
                rngStart.Value = LetzterAdventSonntag(intJahr)
        End Select
        rngStart.NumberFormat = "dd.MM.yyyy"
        Unload Me
    End If
End Sub

Private Sub reJahr_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.reJahr.Value) Then
        Me.lblJahrValue.Caption = Me.reJahr.Value
        intJahr = Me.reJahr.Value
        ShowResult
        Exit Sub
    End If
    
    On Error GoTo Fehler
    Set rngValue = Range(Me.reJahr.Value)
    If rngValue.Cells.Count > 1 Then
      MsgBox "Bitte nur eine Zelle auswählen!"
      Cancel = True
    End If
    
    If IsNumeric(rngValue.Value) = False Then
      MsgBox "Die Zelle muss eine Zahl enthalten."
      Cancel = True
    End If
    
    
    intJahr = rngValue.Value
    Me.lblJahrValue.Caption = intJahr
    ShowResult
    
    Exit Sub
Fehler:
    Select Case Err.Number
        Case 1004
            If Me.reJahr.Value <> vbNullString Then
                MsgBox "Es wurde kein gültiger Bereich eingegeben."
                Err.Clear
                Cancel = True
            End If
            Err.Clear
        Case Else
            MsgBox Err.Number & " - " & Err.Description, , "Fehler"
    End Select
    
End Sub

Private Sub UserForm_Initialize()
    Me.lblJahrValue.Caption = ""
    Me.lblResult.Caption = ""

    Set rngStart = ActiveCell
End Sub

Private Sub ShowResult()
    Select Case strFunction
        Case "Ostern"
            Me.lblResult.Caption = Ostern(intJahr)
        Case "Advent"
            Me.lblResult.Caption = LetzterAdventSonntag(intJahr)
    End Select
End Sub

Public Sub InitForm(ByVal strFunctionDef As String)
    strFunction = strFunctionDef
    Select Case strFunction
        Case "Ostern"
            Me.Caption = "Oster-Funktion"
            Me.lblInfo = "Die Funktion Ostern(Jahr) gibt das Datum des Ostersonnstags für das gegebene Jahr zurück."
        Case "Advent"
            Me.Caption = "LetzterAdventSonntag-Funktion"
            Me.lblInfo = "Die Funktion LetzterAdventSonntag(Jahr) gibt das Datum des 4. Adventsonntags für das gegebene Jahr zurück."
   End Select
End Sub
