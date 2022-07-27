Attribute VB_Name = "mdl_zip"
Option Explicit
Option Private Module

'taken from https://www.rondebruin.nl/win/s7/win001.htm

Private Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, VBA.Chr$(80) & VBA.Chr$(75) & VBA.Chr$(5) & VBA.Chr$(6) & String(18, 0)
    Close #1
End Sub


Private Function bIsBookOpen(ByVal szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Private Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function

Sub Zip_File_Or_Files()
    Dim strDate As String, DefPath As String, sFName As String
    Dim oApp As Object, iCtr As Long, i As Integer
    Dim FName, vArr, FileNameZip
    
    Dim strPath As String
    strPath = AddIns(strVBProjects).path & "\"
    
    FileNameZip = clearPath(strPath & "download_inoHolidays.zip")
    ReDim FName(0)
    FName(0) = clearPath(strPath & "inoHolidays.xlam")
    
    Dim strFile As String
    strFile = dir(strPath & "countrycodes\" & "*.inocsv")
    Do While strFile <> ""
        
        Dim intC As Integer
        intC = UBound(FName) + 1
        ReDim Preserve FName(intC)
        FName(intC) = strPath & "countrycodes\" & strFile
        strFile = dir()
    Loop
    'FName(1) = clearPath(strPath & "countrycodes\de.inocsv")

    If IsArray(FName) = False Then
        'do nothing
    Else
        'Create empty Zip File
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        i = 0
        For iCtr = LBound(FName) To UBound(FName)
            vArr = Split97(FName(iCtr), "\")
            sFName = vArr(UBound(vArr))
            'Copy the file to the compressed folder
            i = i + 1
            oApp.Namespace(FileNameZip).CopyHere FName(iCtr)

            'Keep script waiting until Compressing is done
            On Error Resume Next
            Do Until oApp.Namespace(FileNameZip).items.Count = i
                Application.Wait (Now + TimeValue("0:00:01"))
            Loop
            On Error GoTo 0
        Next iCtr

        'MsgBox "You find the zipfile here: " & FileNameZip
    End If
End Sub

