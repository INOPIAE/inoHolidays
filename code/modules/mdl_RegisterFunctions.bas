Attribute VB_Name = "mdl_RegisterFunctions"
Option Explicit
'https://jkp-ads.com/Articles/RegisterUDF00.asp

Sub RegisterFunction()
   
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

    'Function DayOfMonth
    ReDim vArgDescr(4)
    
    vArgDescr(1) = strRegister(8)
    vArgDescr(2) = strRegister(9)
    vArgDescr(3) = strRegister(10)
    vArgDescr(4) = strRegister(11)
    
    Application.MacroOptions _
        Macro:="DayOfMonth", _
        Description:=strRegister(12), _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr
        
    'Function getIslamicDate
    ReDim vArgDescr(1)
    
    vArgDescr(1) = strRegister(13)

    Application.MacroOptions _
        Macro:="getIslamicDate", _
        Description:=strRegister(14), _
        Category:="inoHolidays", _
        ArgumentDescriptions:=vArgDescr

    'Function getChristianDate
    ReDim vArgDescr(1)
    
    vArgDescr(1) = strRegister(15)

    Application.MacroOptions _
        Macro:="getChristianDate", _
        Description:=strRegister(16) & vbNewLine & strRegister(17), _
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

    'Function DayOfMonth
    ReDim vArgDescr(4)
    Application.MacroOptions _
        Macro:="DayOfMonth", _
        Description:=strRegister(12), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
        
    'Function getIslamicDate
    ReDim vArgDescr(1)
    Application.MacroOptions _
        Macro:="getIslamicDate", _
        Description:=strRegister(14), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr

    'Function getChristianDate
    ReDim vArgDescr(1)
    Application.MacroOptions _
        Macro:="getChristianDate", _
        Description:=strRegister(16) & vbNewLine & strRegister(17), _
        Category:=14, _
        ArgumentDescriptions:=vArgDescr
End Sub

