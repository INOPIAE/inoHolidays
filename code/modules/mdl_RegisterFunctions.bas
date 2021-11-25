Attribute VB_Name = "mdl_RegisterFunctions"
Option Explicit
'https://jkp-ads.com/Articles/RegisterUDF00.asp

Sub RegisterFunction()
    rbSetLanguage
    DetectLanguage
    
    Dim vArgDescr() As Variant
    
    'Function Easter
    ReDim vArgDescr(1)
    
    vArgDescr(1) = strRegister(0)

    Application.MacroOptions _
        Macro:="Easter", _
        Description:=strRegister(1), _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr
        
    'Function LastAdvent
    ReDim vArgDescr(1)
    
    vArgDescr(1) = strRegister(2)

    Application.MacroOptions _
        Macro:="LastAdvent", _
        Description:=strRegister(3), _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr

    'Function isHoliday
    ReDim vArgDescr(3)
    
    vArgDescr(1) = strRegister(4)
    vArgDescr(2) = strRegister(5)
    vArgDescr(3) = strRegister(6)

    Application.MacroOptions _
        Macro:="isHoliday", _
        Description:=strRegister(7), _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr

End Sub

Sub UnRegisterFunction()
    'Make sure the array below has the same size as the original number of arguments
    Dim vArgDescr() As Variant
    
    'Function Easter
    ReDim vArgDescr(1)
    Application.MacroOptions _
        Macro:="Easter", _
        Description:=strRegister(1), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
        
    'Function LastAdvent
    ReDim vArgDescr(1)
    Application.MacroOptions _
        Macro:="LastAdvent", _
        Description:=strRegister(2), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr

    'Function isHoliday
    ReDim vArgDescr(3)
    Application.MacroOptions _
        Macro:="isHoliday", _
        Description:=strRegister(7), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr

End Sub

