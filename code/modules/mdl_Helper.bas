Attribute VB_Name = "mdl_Helper"
Option Explicit
Option Private Module

Public Function mdlHelper_FileExists(ByVal strFilePath As String) As Boolean
      mdlHelper_FileExists = False
      On Error Resume Next
      If Not dir(strFilePath) = vbNullString And Not strFilePath = vbNullString Then
          mdlHelper_FileExists = True
      End If
End Function

Public Function mdlHelper_FolderExists(ByVal strFilePath As String) As Boolean
      mdlHelper_FolderExists = False
      On Error Resume Next
      If Not dir(strFilePath, vbDirectory) = vbNullString And Not strFilePath = vbNullString Then
          mdlHelper_FolderExists = True
      End If
End Function

Public Function clearPath(ByVal strPath) As String
    clearPath = VBA.Replace(strPath, "\\", "\")
End Function

Public Function printF(ByVal strText As String, ParamArray Args()) As String
' © codekabinett.com - You may use, modify, copy, distribute this code as long as this line remains

    Dim i           As Integer
    Dim strRetVal   As String
    Dim startPos    As Integer
    Dim endPos      As Integer
    Dim formatString As String
    Dim argValueLen As Integer
    strRetVal = strText
    
    For i = LBound(Args) To UBound(Args)
        argValueLen = Len(CStr(i))
        startPos = InStr(strRetVal, "{" & CStr(i) & ":")
        If startPos > 0 Then
            endPos = InStr(startPos + 1, strRetVal, "}")
            formatString = Mid(strRetVal, startPos + 2 + argValueLen, endPos - (startPos + 2 + argValueLen))
            strRetVal = Mid(strRetVal, 1, startPos - 1) & VBA.Format(Nz(Args(i), ""), formatString) & Mid(strRetVal, endPos + 1)
        Else
            strRetVal = Replace(strRetVal, "{" & CStr(i) & "}", Nz(Args(i), ""))
        End If
    Next i

    printF = strRetVal

End Function

Public Function Nz(p1, Optional p2) As Variant
    Select Case True
        Case Not IsNull(p1)
            Nz = p1
        Case IsMissing(p2)
            Nz = Empty
        Case Else
            Nz = p2
    End Select
 
End Function



