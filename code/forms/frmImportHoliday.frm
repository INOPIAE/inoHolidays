VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmImportHoliday 
   Caption         =   "Import Holidays"
   ClientHeight    =   4260
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

Public GivenYear As Integer
Private Countries() As String
Public Country As String

Private Sub cmdCancel_Click()
    ImportBln = False
    Unload Me
End Sub

Private Sub cmdImport_Click()
    If IsNumeric(Me.txtJahr.Value) = False Then
        MsgBox t("GivenYear needs to be entered as number."), , strErrorCaptionHint
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    
    Dim rng As Range
    If Me.reZelle.Value = vbNullString Then
        MsgBox t("Select a cell."), , strErrorCaptionHint
        With Me.reZelle
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    Set rng = Range(Me.reZelle.Value)
    If rng.Cells.Count > 1 Then
        MsgBox t("Only one cell must be selected."), , strErrorCaptionHint
        With Me.reZelle
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    
    ImportBln = True
    ImportGivenYear = Me.txtJahr.Value
    ImportCountry = Me.cboCountry.text
    Set ImportRange = rng
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    FillCountries
    InitLanguage
End Sub

Private Sub FillCountries()
    Dim strPath As String
    Dim strFile As String
    Dim intCount As Integer
    strPath = AddIns(strVBProjects).path & "\countrycodes\"
    strFile = dir(strPath & "*")
    Do While strFile <> ""
        ReDim Preserve Countries(intCount)
        Dim c() As String
        c = Split(strFile, ".")
        If c(1) = "inocsv" Then
            Me.cboCountry.AddItem (c(0))
            intCount = intCount + 1
        End If
        strFile = dir()
    Loop
End Sub

Public Sub FillForm(ByVal j As Integer, ByVal c As String)
    GivenYear = j
    Country = c
    Me.txtJahr.text = GivenYear
    Me.cboCountry.text = Country
End Sub

Private Sub InitLanguage()
    TranslateForm Me
    Me.lblInfo.Caption = t("There should be 3 emtpy columns and 20 empty rows starting from the given cell. Filled cell will be overwriten without further notification.")
End Sub
