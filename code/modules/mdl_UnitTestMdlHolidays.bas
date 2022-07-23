Attribute VB_Name = "mdl_UnitTestMdlHolidays"
Option Explicit
Option Private Module

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
Private Sub TestDayOfMonth()                       'TODO Test umbenennen
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim GivenYear As Integer
    Dim GivenMonth As Integer
    Dim DayOfWeek As Integer
    Dim NumOfMonth As Integer
    Dim TestDate As Variant
    'Ausfuehren:
    GivenYear = 2020
    GivenMonth = 1
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 1)
    If TestDate <> #1/6/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 2)
    If TestDate <> #1/13/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 3)
    If TestDate <> #1/20/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 4)
    If TestDate <> #1/27/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 5)
    If TestDate <> CVErr(xlErrNA) Then GoTo TestFehlschlag

    TestDate = DayOfMonth(GivenYear, GivenMonth, vbMonday, 6)
    If TestDate <> #1/27/2020# Then GoTo TestFehlschlag
    
    
    GivenMonth = 2
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 1)
    If TestDate <> #2/4/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 2)
    If TestDate <> #2/11/2020# Then GoTo TestFehlschlag

    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 3)
    If TestDate <> #2/18/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 4)
    If TestDate <> #2/25/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 5)
    If TestDate <> CVErr(xlErrNA) Then GoTo TestFehlschlag

    TestDate = DayOfMonth(GivenYear, GivenMonth, vbSaturday, 5)
    If TestDate <> #2/29/2020# Then GoTo TestFehlschlag
    
    TestDate = DayOfMonth(GivenYear, GivenMonth, vbTuesday, 6)
    If TestDate <> #2/25/2020# Then GoTo TestFehlschlag
    'Validieren:
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
    Resume TestEnde
End Sub

