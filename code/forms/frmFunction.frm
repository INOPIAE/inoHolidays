VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFunction 
   Caption         =   "no translation"
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

Private Const strDateFormat = "MM/dd/yyyy" ' < translate


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
        rngStart.NumberFormat = strDateFormat
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
        rngStart.NumberFormat = strDateFormat
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
      MsgBox t("Only one cell must be selected."), , strErrorCaptionHint
      Cancel = True
    End If
    
    If IsNumeric(rngValue.Value) = False Then
      MsgBox t("The cell must contain a number."), , strErrorCaptionHint
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
                MsgBox t("No valid range entered."), , strErrorCaptionHint
                Err.Clear
                Cancel = True
            End If
            Err.Clear
        Case Else
            MsgBox Err.Number & " - " & Err.Description, , strErrorCaption
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
            Me.lblResult.Caption = Easter(intGivenYear)
        Case "Advent"
            Me.lblResult.Caption = LastAdvent(intGivenYear)
    End Select
End Sub

Public Sub InitForm(ByVal strFunctionDef As String)
    strFunction = strFunctionDef
    TranslateForm Me
    Select Case strFunction
        Case "Ostern"
            Me.Caption = t("Function {}", "Easter")
            Me.lblInfo = t("The function {} returns the date of Easter Sunday of the given year.", "Easter(GivenYear)")
        Case "Advent"
            Me.Caption = t("Function {}", "LastAdvent")
            Me.lblInfo = t("The function {} returns the date of 4th Advent Sunday of the given year.", "LastAdvent(GivenYear)")
   End Select
End Sub

