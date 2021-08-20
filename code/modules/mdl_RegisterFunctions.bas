Attribute VB_Name = "mdl_RegisterFunctions"
Option Explicit
'https://jkp-ads.com/Articles/RegisterUDF00.asp

Sub RegisterFunction()
    Dim vArgDescr() As Variant
    
    'Function Ostern
    ReDim vArgDescr(1)
    
    vArgDescr(1) = "Jahr - Das Jahr f�r den Ostersonntag"

    Application.MacroOptions _
        Macro:="Ostern", _
        Description:="Gibt das Datum des Ostersonntags des angegeben Jahres zur�ck.", _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr
End Sub

Sub UnRegisterFunction()
    'Make sure the array below has the same size as the original number of arguments
    Dim vArgDescr() As Variant
    
    'Function Ostern
    ReDim vArgDescr(1)
    Application.MacroOptions _
        Macro:="Ostern", _
        Description:="Gibt das Datum des Ostersonntags des angegeben Jahres zur�ck.", _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
End Sub

