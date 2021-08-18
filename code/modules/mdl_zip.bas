Attribute VB_Name = "mdl_zip"
Option Explicit

'taken from https://www.rondebruin.nl/win/s7/win001.htm

Sub NewZip(sPath)
'Create empty Zip File
'Changed by keepITcool Dec-12-2005
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub


Function bIsBookOpen(ByVal szBookName As String) As Boolean
' Rob Bovey
    On Error Resume Next
    bIsBookOpen = Not (Application.Workbooks(szBookName) Is Nothing)
End Function


Function Split97(sStr As Variant, sdelim As String) As Variant
'Tom Ogilvy
    Split97 = Evaluate("{""" & _
                       Application.Substitute(sStr, sdelim, """,""") & """}")
End Function

Sub Zip_File_Or_Files()
    Dim strDate As String, DefPath As String, sFName As String
    Dim oApp As Object, iCtr As Long, I As Integer
    Dim FName, vArr, FileNameZip
    
    Dim strPath As String
    strPath = AddIns(strVBProjects).Path & "\"
    
    FileNameZip = clearPath(strPath & "download_inoHolidays.zip")
    ReDim FName(1)
    FName(0) = clearPath(strPath & "inoHolidays.xlam")
    FName(1) = clearPath(strPath & "countrycodes\de.inocsv")

    If IsArray(FName) = False Then
        'do nothing
    Else
        'Create empty Zip File
        NewZip (FileNameZip)
        Set oApp = CreateObject("Shell.Application")
        I = 0
        For iCtr = LBound(FName) To UBound(FName)
            vArr = Split97(FName(iCtr), "\")
            sFName = vArr(UBound(vArr))
            'Copy the file to the compressed folder
            I = I + 1
            oApp.Namespace(FileNameZip).CopyHere FName(iCtr)

            'Keep script waiting until Compressing is done
            On Error Resume Next
            Do Until oApp.Namespace(FileNameZip).items.Count = I
                Application.Wait (Now + TimeValue("0:00:01"))
            Loop
            On Error GoTo 0
        Next iCtr

        'MsgBox "You find the zipfile here: " & FileNameZip
    End If
End Sub

