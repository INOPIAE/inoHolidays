Attribute VB_Name = "mdl_UnitTestMdlIslamicCalender"
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
Private Sub TestGetIslamicDate()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim ChristianDate As Date
    Dim ResultDate As String

    Dim TestDate As Variant
    'Ausfuehren:
    
    ' 1444 AH / 1445 AH
    ResultDate = "29. Dhu'l-Hijja 1444 AH"
    ChristianDate = #7/18/2023#
    
    TestDate = getIslamicDate(ChristianDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    ResultDate = "1. Muharram 1445 AH"
    ChristianDate = #7/19/2023#
    
    TestDate = getIslamicDate(ChristianDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    ' 1444 AH / 1445 AH
    
    ResultDate = "29. Dhu'l-Hijja 1445 AH"
    ChristianDate = #7/6/2024#
    
    TestDate = getIslamicDate(ChristianDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    ResultDate = "30. Dhu'l-Hijja 1445 AH"
    ChristianDate = #7/7/2024#
    
    TestDate = getIslamicDate(ChristianDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    ResultDate = "1. Muharram 1446 AH"
    ChristianDate = #7/8/2024#
    
    TestDate = getIslamicDate(ChristianDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    'Validieren:
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description _
        & "Resulttest: " & ResultDate
    Resume TestEnde
End Sub

'@TestMethod("Uncategorized")
Private Sub TestGetChristianDate()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim IslamicDate As String
    Dim ResultDate As Date

    Dim TestDate As Variant
    'Ausfuehren:
    
    ' 1444 AH / 1445 AH
    IslamicDate = "29. Dhu'l-Hijja 1444 AH"
    ResultDate = #7/18/2023#
    
    TestDate = getChristianDate(IslamicDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    IslamicDate = "1. Muharram 1445 AH"
    ResultDate = #7/19/2023#
    
    TestDate = getChristianDate(IslamicDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    ' 1444 AH / 1445 AH
    
    IslamicDate = "29. Dhu'l-Hijja 1445 AH"
    ResultDate = #7/6/2024#
    
    TestDate = getChristianDate(IslamicDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    IslamicDate = "30. Dhu'l-Hijja 1445 AH"
    ResultDate = #7/7/2024#
    
    TestDate = getChristianDate(IslamicDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    IslamicDate = "1. Muharram 1446 AH"
    ResultDate = #7/8/2024#
    
    TestDate = getChristianDate(IslamicDate)
    If TestDate <> ResultDate Then GoTo TestFehlschlag
    
    'Validieren:
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description
    Resume TestEnde
End Sub


'@TestMethod("Uncategorized")
Private Sub TestGetIslamicDateCont()
    On Error GoTo TestFehlschlag
    
    'Einrichten:
    Dim ChristianDate As Date
    Dim ResultDate As String
    Dim intC As Integer
    Dim TestDate As Variant
    'Ausfuehren:
    
    ChristianDate = #8/1/2019#
    
    For intC = 1 To 10700

        ChristianDate = DateAdd("d", 1, ChristianDate)
    
        ResultDate = getIslamicDate(ChristianDate)
        TestDate = getChristianDate(ResultDate)
        If TestDate <> ChristianDate Then GoTo TestFehlschlag
        
        
    Next
    
    'Validieren:
    Assert.Succeed

TestEnde:
    Exit Sub
TestFehlschlag:
    
    Assert.fail "Test hat diesen Fehler ergeben: #" & Err.Number & " - " & Err.Description _
        & "Testcase: " & ChristianDate
    Resume TestEnde
End Sub


