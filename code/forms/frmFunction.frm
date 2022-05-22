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
1   Unload Me
End Sub

Private Sub cmdImport_Click()
2   If intGivenYear > 0 Then
3       Select Case strFunction
            Case "Ostern"
4               rngStart.Formula = "=Easter(" & Me.reGivenYear.Value & ")"
5           Case "Advent"
6               rngStart.Formula = "=LastAdvent(" & Me.reGivenYear.Value & ")"
7       End Select
8       rngStart.NumberFormat = strDateFormat
9       Unload Me
10  End If
End Sub

Private Sub cmdValue_Click()
11  If intGivenYear > 0 Then
12      Select Case strFunction
            Case "Ostern"
13              rngStart.Value = Easter(intGivenYear)
14          Case "Advent"
15              rngStart.Value = LastAdvent(intGivenYear)
16      End Select
17      rngStart.NumberFormat = strDateFormat
18      Unload Me
19  End If
End Sub

Private Sub reGivenYear_Exit(ByVal Cancel As MSForms.ReturnBoolean)
20  If IsNumeric(Me.reGivenYear.Value) Then
21      Me.lblGivenYearValue.Caption = Me.reGivenYear.Value
22      intGivenYear = Me.reGivenYear.Value
23      ShowResult
24      Exit Sub
25  End If
    
26  On Error GoTo Fehler
27  Set rngValue = Range(Me.reGivenYear.Value)
28  If rngValue.Cells.Count > 1 Then
29    MsgBox t("Only one cell must be selected."), , strErrorCaptionHint
30    Cancel = True
31  End If
    
32  If IsNumeric(rngValue.Value) = False Then
33    MsgBox t("The cell must contain a number."), , strErrorCaptionHint
34    Cancel = True
35  End If
    
    
36  intGivenYear = rngValue.Value
37  Me.lblJahrValue.Caption = intGivenYear
38  ShowResult
    
39  Exit Sub
40 Fehler:
41  Select Case Err.Number
        Case 1004
42          If Me.reGivenYear.Value <> vbNullString Then
43              MsgBox t("No valid range entered."), , strErrorCaptionHint
44              Err.Clear
45              Cancel = True
46          End If
47          Err.Clear
48      Case Else
49          MsgBox t("The error {} occured in line {}:", Err.Number, Erl) & vbNewLine & Err.Description, , strErrorCaption
50  End Select
    
End Sub

Private Sub UserForm_Initialize()
51  Me.lblJahrValue.Caption = ""
52  Me.lblResult.Caption = ""

53  Set rngStart = ActiveCell
End Sub

Private Sub ShowResult()
54  Select Case strFunction
        Case "Ostern"
55          Me.lblResult.Caption = Easter(intGivenYear)
56      Case "Advent"
57          Me.lblResult.Caption = LastAdvent(intGivenYear)
58  End Select
End Sub

Public Sub InitForm(ByVal strFunctionDef As String)
59  strFunction = strFunctionDef
60  TranslateForm Me
61  Select Case strFunction
        Case "Ostern"
62          Me.Caption = t("Function {}", "Easter")
63          Me.lblInfo = t("The function {} returns the date of Easter Sunday of the given year.", "Easter(GivenYear)")
64      Case "Advent"
65          Me.Caption = t("Function {}", "LastAdvent")
66          Me.lblInfo = t("The function {} returns the date of 4th Advent Sunday of the given year.", "LastAdvent(GivenYear)")
67 End Select
End Sub

