Attribute VB_Name = "mdl_Oultlook"
Option Explicit

Private olApp As Outlook.Application
Private olHoliday As Outlook.AppointmentItem
Private olHolidaytest As Outlook.AppointmentItem

Sub EnterHoliday(holiday As String, myDate As Date, Location As String, Optional status As Integer = 0)

    If olApp Is Nothing Then
        Set olApp = GetObject(, "Outlook.Application")
    End If
    
    Set olHoliday = olApp.CreateItem(olAppointmentItem)
    
    With olHoliday
        .AllDayEvent = True
        .Start = Format(myDate, "dd.mm.yyyy")
    
        .Subject = holiday
        .ReminderSet = False
        .Location = Location
        .BusyStatus = status  'free = 0, busy =2
'TODO Language
        .Categories = "Feiertag, inoHolidays"
        .Save

        '.Display
        
     End With
    
     Set olHoliday = Nothing

End Sub

Public Function getHoliday(myDate As Date, holiday As String, Location As String) As Outlook.AppointmentItem
    Dim myStart As Date
    Dim myEnd As Date
    Dim olCalendar As Outlook.Folder
    Dim olItems As Outlook.items
    Dim olResItems As Outlook.items
    Dim olAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    
    Set getHoliday = Nothing
    
    myStart = myDate
    myEnd = DateAdd("d", 1, myStart)
    
    If olApp Is Nothing Then
        Set olApp = GetObject(, "Outlook.Application")
    End If

    Set olCalendar = olApp.Session.GetDefaultFolder(olFolderCalendar)
    Set olItems = olCalendar.items
     
    olItems.IncludeRecurrences = True
    olItems.Sort "[Start]"
     
    strRestriction = "[Start] <= '" & Format$(myEnd, "DD.MM.YYYY hh:mm AMPM") _
    & "' AND [End] >= '" & Format(myStart, "DD.MM.YYYY hh:mm AMPM") & "' and [Location] ='" & Location & "'"

    Set olResItems = olItems.Restrict(strRestriction)
     
    For Each olAppt In olResItems
        If olAppt.Categories Like "*inoHolidays*" And olAppt.Subject = holiday Then
            Set getHoliday = olAppt
            Exit Function
        End If
    Next
End Function


Public Sub ImportOutlookHolidays(ByVal myYear As Integer, ByVal CountryDef As String, ByVal StateDef As String, ByVal blnBusy As Boolean)
    Dim clsH As New clsHolidays
    Dim arr
    Dim i As Integer
    Dim holiday As AppointmentItem
    Dim intNew As Integer
    Dim intChanged As Integer
    Dim intNoChange As Integer
    Dim Country As String
    
    On Error GoTo MyError
    
    If olApp Is Nothing Then
        Set olApp = GetObject(, "Outlook.Application")
    End If
    
    arr = clsH.GetAllHolidays(myYear, CountryDef)
    
    Country = getLocation(CountryDef)
    
    For i = 0 To UBound(arr)
        Dim h() As String
        Dim intBusy As Integer
        h = Split(arr(i), ";")
        Set holiday = getHoliday(CDate(h(0)), Trim(h(1)) & ", " & Trim(h(2)), Country)
        If h(2) Like "*" & StateDef & "*" Or Trim(h(2)) = "All" Then
            intBusy = olBusy
        Else
            intBusy = olFree
        End If
        If blnBusy = False Then
            intBusy = olFree
        End If
        If holiday Is Nothing Then
            EnterHoliday Trim(h(1)) & ", " & Trim(h(2)), CDate(h(0)), Country, intBusy
            intNew = intNew + 1
        Else
            If holiday.BusyStatus <> intBusy Then
                holiday.BusyStatus = intBusy
                intChanged = intChanged + 1
            Else
                intNoChange = intNoChange + 1
            End If
        End If
    Next
    
    MsgBox t("{} entries processed for {}. Thereof \n" _
        & "{} new entries\n" _
        & "{} changed entries\n" _
        & "{} unchanged entries", intNew + intChanged + intNoChange, myYear, intNew, intChanged, intNoChange)
    
    Exit Sub
    
MyError:
    Select Case Err.Number
        Case 429
            MsgBox t("Outlook must be started."), , strErrorCaptionHint
        Case Else
            MsgBox Err.Number & " " & Err.Description, , strErrorCaption
    End Select
End Sub

Public Sub deleteHolidaysYear(myYear As Integer, Location As String)
    Dim myStart As Date
    Dim myEnd As Date
    Dim olCalendar As Outlook.Folder
    Dim olItems As Outlook.items
    Dim olResItems As Outlook.items
    Dim olAppt As Outlook.AppointmentItem
    Dim strRestriction As String
    
    Dim Country As String
    
    On Error GoTo MyError
    
    If olApp Is Nothing Then
        Set olApp = GetObject(, "Outlook.Application")
    End If


    arr = clsH.GetAllHolidays(myYear, CountryDef)
    
    Country = getLocation(Location)

    myStart = DateSerial(myYear, 1, 1)
    myEnd = DateSerial(myYear + 1, 1, 1)

    Set olCalendar = olApp.Session.GetDefaultFolder(olFolderCalendar)
    Set olItems = olCalendar.items
     
    olItems.IncludeRecurrences = True
    olItems.Sort "[Start]"
     
    strRestriction = "[Start] < '" & Format$(myEnd, "DD.MM.YYYY hh:mm AMPM") _
    & "' AND [End] >= '" & Format(myStart, "DD.MM.YYYY hh:mm AMPM") & "' and [Location] ='" & Location & "'"

    Set olResItems = olItems.Restrict(strRestriction)
    Dim intCount As Integer
    For Each olAppt In olResItems
        If olAppt.Categories Like "*inoHolidays*" Then
            olAppt.Delete
            intCount = intCount + 1
        End If
    Next
    MsgBox t("{} entries deleted for {}.", intCount, myYear)
    
    Exit Sub
    
MyError:
    Select Case Err.Number
        Case 429
            MsgBox t("Outlook must be started."), , strErrorCaptionHint
        Case Else
            MsgBox Err.Number & " " & Err.Description, , strErrorCaption
    End Select
End Sub

Public Sub deleteHolidays()
    Dim myStart As Date
    Dim myEnd As Date
    Dim olCalendar As Outlook.Folder
    Dim olItems As Outlook.items
    Dim olResItems As Outlook.items
    Dim olAppt As Outlook.AppointmentItem
    Dim strRestriction As String
  
    On Error GoTo MyError
    
    If olApp Is Nothing Then
        Set olApp = GetObject(, "Outlook.Application")
    End If

    Set olCalendar = olApp.Session.GetDefaultFolder(olFolderCalendar)
    Set olItems = olCalendar.items
     
    olItems.IncludeRecurrences = True
    olItems.Sort "[Start]"

    Dim intCount As Integer
    For Each olAppt In olItems
        If olAppt.Categories Like "*inoHolidays*" Then
            olAppt.Delete
            intCount = intCount + 1
        End If
    Next
    
    MsgBox t("{} entries deleted.", intCount)
    
    Exit Sub
        
MyError:
    Select Case Err.Number
        Case 429
            MsgBox t("Outlook must be started."), , strErrorCaptionHint
        Case Else
            MsgBox Err.Number & " " & Err.Description, , strErrorCaption
    End Select
End Sub

Public Function getLocation(ByVal strLocation As String) As String
    Select Case strLocation
        Case "de"
            getLocation = "Deutschland"
        Case "at"
            getLocation = "Österreich"
        Case "ch"
            getLocation = "Schweiz"
    End Select
End Function
