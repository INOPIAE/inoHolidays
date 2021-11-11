VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmOutlookImport 
   Caption         =   "Import Feiertage nach Outlook"
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
        MsgBox strFrmHolidays(5)
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboCountry.Value = vbNullString Then
        MsgBox strFrmHolidays(5)
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
        MsgBox strFrmOutlook(6)
        With Me.txtJahr
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboCountry.Value = vbNullString Then
        MsgBox strFrmOutlook(7)
        With Me.cboCountry
            .SetFocus
            .SelStart = 0
            .SelLength = 100
        End With
        Exit Sub
    End If
    If Me.cboState.Value = vbNullString Then
        MsgBox strFrmOutlook(8)
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
    strPath = AddIns(strVBProjects).Path & "\countrycodes\"
    strFile = Dir(strPath & "*")
    Do While strFile <> ""
        ReDim Preserve Countries(intCount)
        Dim c() As String
        c = Split(strFile, ".")
        Me.cboCountry.AddItem (c(0))
        intCount = intCount + 1
        strFile = Dir()
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
    If lc = 0 Then SetLanguage
    Me.Caption = strFrmOutlook(0)
    Me.lblJahr.Caption = strFrmOutlook(1)
    Me.lblCountry.Caption = strFrmOutlook(2)
    Me.lblState.Caption = strFrmOutlook(3)
    Me.chkBusy.Caption = strFrmOutlook(4)
    Me.lblInfo.Caption = strFrmOutlook(5)
    Me.cmdCancel.Caption = strCmd(1)
    Me.cmdImport.Caption = strCmd(5)
    Me.cmdDeleteYear.Caption = strCmd(6)
    Me.cmdDeleteAll.Caption = strCmd(7)
End Sub
