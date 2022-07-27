Attribute VB_Name = "mdl_Project"
Option Explicit

' entnommen https://answers.microsoft.com/de-de/msoffice/forum/all/ms-project-2016-%C3%B6ffnen-via-vba-makro-im-excel/0212097f-3f70-401e-8bb5-8423a6c1e002
'******************************************************************************
'******************************************************************************
'Globale Konstanten
Global Const c_strGUID = "{A7107640-94DF-1068-855E-00DD01075445}," ' MSProject

Sub SetReferences()
    Dim strGUID As Variant
    Dim theRef As Variant
    Dim i, v_anzahl As Long
    
    '*****************************************************************************
    '**** Referenzen setzen, um Outlook und Excel zu aktivieren
    '*****************************************************************************
        
    '****Referencen als Konstanten definieren
    strGUID = Split(c_strGUID, ",")
    v_anzahl = UBound(strGUID)
    
    
    '****Hier ausdrücklich keine explizite Fehlerbehandlung
        On Error Resume Next
    '****Fehlerhafte Referenzen entfernen
        For i = Application.VBE.ActiveVBProject.references.Count To 1 Step -1
            Set theRef = Application.VBE.ActiveVBProject.references.Item(i)
            If theRef.IsBroken = True Then
                Application.VBE.ActiveVBProject.references.Remove theRef
            End If
        Next i
       
    '****Unter c_strGUID definierte Referencen setzen
        For i = 0 To v_anzahl - 1
             'Fehler löschen, um expizite Fehlerbewertung zu ermöglichen
            Err.Clear
            
             'Referenz setzen
            Application.VBE.ActiveVBProject.references.AddFromGuid GUID:=strGUID(i), Major:=1, Minor:=0
             'Fehler interpretieren
            Select Case Err.Number
            Case 32813
                 'Referenz schon gesetzt - keine Aktivität erforderlich
            Case vbNullString
                 'Referenz ohne Problem gesetzt
            Case Else
                 'Unbekannter Fehler - Abbruch
                 GoTo Ref_Error
            End Select
        Next i
    '****Unterdrücken der Fehlerbehandlung wird wieder deaktiviert
    On Error GoTo 0
    Exit Sub
Ref_Error:
    MsgBox "Referenz nicht gesetzt, Fehler: " & Err.Number & " - " & Err.Description
End Sub


Sub ImportProjectHolidays(ByVal myYear As Integer, ByVal CountryDef As String, ByVal StateDef As String, _
    ByVal Calendarname As String, ByVal pjFileE, pjAppE, _
    ByVal myYearTo As String)
    
    Dim Cal As MSProject.Calendar
    Dim clsH As New clsHolidays
    Dim arr
    Dim blnCal As Boolean
    Dim intNew As Integer
    Dim intChanged As Integer
    Dim intNoChange As Integer
    Dim Country As String
    Dim i As Integer
    Dim intYear As Integer
    
    On Error GoTo Fehler
    
    Dim pjApp As MSProject.Application
    Dim pjFile As MSProject.Project
    
    Set pjApp = pjAppE
    Set pjFile = pjFileE
    
    For Each Cal In pjFile.BaseCalendars
        If Cal.Name = Calendarname Then
            blnCal = True
            Exit For
        End If
    Next
    
    
    If blnCal = False Then
        pjApp.BaseCalendarCreate Name:=Calendarname
    End If
    If IsNumeric(myYearTo) = False Or VBA.Trim(myYearTo) = vbNullString Then
        myYearTo = myYear
    End If
    For intYear = myYear To myYearTo
    
        arr = clsH.GetAllHolidays(intYear, CountryDef)
    
        For i = 0 To UBound(arr)
            Dim h() As String
            Dim intBusy As Integer
            h = Split(arr(i), ";")
            

            If h(2) Like "*" & StateDef & "*" Or VBA.Trim(h(2)) = "All" Then
                pjFile.BaseCalendars(Calendarname).Exceptions.Add Type:=1, Start:=CDate(h(0)), Finish:=CDate(h(0)), Name:=VBA.Trim(h(1))
                intNew = intNew + 1
Weiter:
            End If
        Next
    Next
    MsgBox t("{} entries processed for {}. Thereof \n" _
        & "{} new entries\n" _
        & "{} changed entries\n" _
        & "{} unchanged entries", intNew + intChanged + intNoChange, myYear, intNew, intChanged, intNoChange)
    
    Exit Sub
Fehler:
    Select Case Err.Number
        Case 1101
            Err.Clear
            intNoChange = intNoChange + 1
            Resume Weiter
        Case Else
            MsgBox Err.Number & " " & Err.Description
    End Select
End Sub

Sub ShowProjectImport()
    Dim ProFilename As String
    With Application.FileDialog(msoFileDialogOpen)
        .Filters.Clear
        .Filters.Add "Project file", "*.mpp"
        .Title = "Choose Project file"
        If .Show = -1 Then
            ProFilename = .SelectedItems(1)
        Else
            Exit Sub
        End If
    End With
    
    With frmProjectImport
        .DefineProjectFile ProFilename
       .Show
    End With
End Sub
