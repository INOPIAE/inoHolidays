VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutlookImport 
   Caption         =   "Add Holidays To Outlook"
   ClientHeight    =   4680
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6960
   OleObjectBlob   =   "frmOutlookImport.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmOutlookImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public GivenYear As Integer
Private Countries() As String
Public Country As String

Private Sub cboCountry_Change()
    FillState
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDeleteAll_Click()
    mdl_Oultlook.deleteHolidays
End Sub

Private Sub cmdDeleteYear_Click()
    If IsNumeric(Me.txtJahr.Value) = False Then
        MsgBox t("GivenYear needs to be entered as number."), , strErrorCaptionHint
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboCountry.Value = vbNullString Then
        MsgBox t("Country needs to be given."), , strErrorCaptionHint
        With Me.cboCountry
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    mdl_Oultlook.deleteHolidaysYear Me.txtJahr.Value, Me.cboCountry.Value
End Sub

Private Sub cmdImport_Click()
    If IsNumeric(Me.txtJahr.Value) = False Then
        MsgBox t("Year must be entered as number."), , strErrorCaptionHint
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboCountry.Value = vbNullString Then
        MsgBox t("A country must be selected."), , strErrorCaptionHint
        With Me.cboCountry
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboState.Value = vbNullString Then
        MsgBox t("A state must be selected."), , strErrorCaptionHint
        With Me.cboState
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    mdl_Oultlook.ImportOutlookHolidays Me.txtJahr.Value, Me.cboCountry.Value, Me.cboState.Value, Me.chkBusy.Value
End Sub

Private Sub UserForm_Initialize()
    FillCountries
    FillState
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
        Me.cboCountry.AddItem (c(0))
        intCount = intCount + 1
        strFile = dir()
    Loop
End Sub

Private Sub FillState()
    Me.cboState.Clear
    Dim wkb As Workbook
    Dim wks As Worksheet
    Dim loStates As ListObject
    Dim i As Integer
    
    Set wkb = Application.Workbooks(strVBProjects & ".xlam")
    Set wks = wkb.Worksheets("Konfig")
    Set loStates = wks.ListObjects("Bundeslaender")
    
    For i = 1 To loStates.DataBodyRange.Rows.Count
        If loStates.DataBodyRange.Cells(i, 1).Value = Me.cboCountry.Value Then
            Me.cboState.AddItem (loStates.DataBodyRange.Cells(i, 2).Value)
        End If
    Next
    If Me.cboState.ListCount > 0 Then
        Me.cboState.ListIndex = 0
    End If
End Sub

Private Sub InitLanguage()
    TranslateForm Me
    Me.lblInfo.Caption = t("The first 3 fields must be used.{}If 'Add holiday as busy' is ticked the public holidays will be add as busy for the given State/Region. If State/Region = All only the countrywide holidays are set as busy.", vbNewLine)
End Sub


