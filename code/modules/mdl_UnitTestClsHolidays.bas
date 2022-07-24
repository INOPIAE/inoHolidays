Attribute VB_Name = "mdl_UnitTestClsHolidays"
Option Explicit
Option Private Module

' to use this module the COM add-in rubberduck needs be installed
' https://github.com/rubberduck-vba/Rubberduck

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Private Sub ModulInitialisierung()
    'Diese Methode wird einmal pro Modul ausgeführt.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Private Sub ModulTerminierung()
    'Diese Methode wird einmal pro Modul ausgeführt.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Private Sub TestInitialisierung()
    'Diese Methode wird vor jedem Test ausgeführt..
End Sub

'@TestCleanup
Private Sub TestTerminierung()
    'Diese Methode wird nach jedem Test ausgeführt.
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetHolidayNameFixed()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim h As New clsHolidays
    Dim GivenDate As Date
    Dim Feiertag As String
    
    'Ausfuehren:
    GivenDate = #1/1/2020#
    
    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> "Neujahr" Then GoTo TestFehlschlag
        
    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> "Neujahr" Then GoTo TestFehlschlag
        
    GivenDate = #1/2/2020#
    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> vbNullString Then GoTo TestFehlschlag

    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> vbNullString Then GoTo TestFehlschlag
        
    GivenDate = #3/8/2020#
    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> vbNullString Then GoTo TestFehlschlag

    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> "Int. Frauentag" Then GoTo TestFehlschlag
        
    'Validieren:
    
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetHolidayNameDefault()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim h As New clsHolidays
    Dim GivenDate As Date
    Dim Feiertag As String
    
    'Ausfuehren:
    GivenDate = DateSerial(Year(Now), 1, 1)
    
    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> "Neujahr" Then GoTo TestFehlschlag
        
    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> "Neujahr" Then GoTo TestFehlschlag
        
    GivenDate = DateSerial(Year(Now), 1, 2)

    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> vbNullString Then GoTo TestFehlschlag

    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> vbNullString Then GoTo TestFehlschlag
        
    GivenDate = DateSerial(Year(Now), 3, 8)

    Feiertag = h.GetHolidayName(GivenDate)
    If Feiertag <> vbNullString Then GoTo TestFehlschlag

    Feiertag = h.GetHolidayName(GivenDate, "BE")
    If Feiertag <> "Int. Frauentag" Then GoTo TestFehlschlag
        
    'Validieren:
    
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestIsHolidayFixed()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim h As New clsHolidays
    Dim GivenDate As Date
    
    'Ausfuehren:
    GivenDate = #1/1/2020#
    
    If h.isHoliday(GivenDate) = False Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = False Then GoTo TestFehlschlag
        
    GivenDate = #1/2/2020#
    If h.isHoliday(GivenDate) = True Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = True Then GoTo TestFehlschlag
        
    GivenDate = #3/8/2020#
    If h.isHoliday(GivenDate) = True Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = False Then GoTo TestFehlschlag
        
    'Validieren:
    
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Uncategorized")
Private Sub TestIsHolidayDefault()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim h As New clsHolidays
    Dim GivenDate As Date
    
    'Ausfuehren:
    GivenDate = DateSerial(Year(Now), 1, 1)
    If h.isHoliday(GivenDate) = False Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = False Then GoTo TestFehlschlag
        
    GivenDate = DateSerial(Year(Now), 1, 2)
    If h.isHoliday(GivenDate) = True Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = True Then GoTo TestFehlschlag
        
    GivenDate = DateSerial(Year(Now), 3, 8)
    If h.isHoliday(GivenDate) = True Then GoTo TestFehlschlag
    If h.isHoliday(GivenDate, , "BE") = False Then GoTo TestFehlschlag
    
    'Validieren:
    
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
End Sub
