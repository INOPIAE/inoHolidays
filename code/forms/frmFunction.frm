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
Private intGivenYear As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImport_Click()
    If intGivenYear > 0 Then
        Select Case strFunction
            Case "Ostern"
                rngStart.Formula = "=Easter(" & Me.reGivenYear.Value & ")"
            Case "Advent"
                rngStart.Formula = "=LastAdvent(" & Me.reGivenYear.Value & ")"
        End Select
        rngStart.NumberFormat = strFrmFunction(5)
        Unload Me
    End If
End Sub

Private Sub cmdValue_Click()
    If intGivenYear > 0 Then
        Select Case strFunction
            Case "Ostern"
                rngStart.Value = Easter(intGivenYear)
            Case "Advent"
                rngStart.Value = LastAdvent(intGivenYear)
        End Select
        rngStart.NumberFormat = strFrmFunction(5)
        Unload Me
    End If
End Sub

Private Sub reGivenYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If IsNumeric(Me.reGivenYear.Value) Then
        Me.lblGivenYearValue.Caption = Me.reGivenYear.Value
        intGivenYear = Me.reGivenYear.Value
        ShowResult
        Exit Sub
    End If
    
    On Error GoTo Fehler
    Set rngValue = Range(Me.reGivenYear.Value)
    If rngValue.Cells.Count > 1 Then
      MsgBox strFrmFunction(6)
      Cancel = True
    End If
    
    If IsNumeric(rngValue.Value) = False Then
      MsgBox strFrmFunction(7)
      Cancel = True
    End If
    
    
    intGivenYear = rngValue.Value
    Me.lblJahrValue.Caption = intGivenYear
    ShowResult
    
    Exit Sub
Fehler:
    Select Case Err.Number
        Case 1004
            If Me.reGivenYear.Value <> vbNullString Then
                MsgBox strFrmFunction(8)
                Err.Clear
                Cancel = True
            End If
            Err.Clear
        Case Else
            MsgBox Err.Number & " - " & Err.Description, , strError(0)
    End Select
    
End Sub

Private Sub lblJahr_Click()

End Sub

Private Sub UserForm_Initialize()
    Me.lblJahrValue.Caption = ""
    Me.lblResult.Caption = ""

    Set rngStart = ActiveCell
End Sub

Private Sub ShowResult()
    Select Case strFunction
        Case "Ostern"
            Me.lblResult.Caption = Easter(intGivenYear)
        Case "Advent"
            Me.lblResult.Caption = LastAdvent(intGivenYear)
    End Select
End Sub

Public Sub InitForm(ByVal strFunctionDef As String)
    If lc = 0 Then SetLanguage
    strFunction = strFunctionDef
    Select Case strFunction
        Case "Ostern"
            Me.Caption = strFrmFunction(1)
            Me.lblInfo = strFrmFunction(2)
        Case "Advent"
            Me.Caption = strFrmFunction(3)
            Me.lblInfo = strFrmFunction(4)
   End Select
   Me.cmdCancel.Caption = strCmd(1)
   Me.cmdValue.Caption = strCmd(3)
   Me.cmdImport.Caption = strCmd(4)
   Me.lblJahr.Caption = strFrmFunction(0)
End Sub
