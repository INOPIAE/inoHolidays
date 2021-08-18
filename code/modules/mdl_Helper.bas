Attribute VB_Name = "mdl_Helper"
Option Explicit
Option Private Module

Public Function mdlHelper_FileExists(ByVal strFilePath As String) As Boolean
      mdlHelper_FileExists = False
      On Error Resume Next
      If Not Dir(strFilePath) = vbNullString And Not strFilePath = vbNullString Then
          mdlHelper_FileExists = True
      End If
End Function

Public Function mdlHelper_FolderExists(ByVal strFilePath As String) As Boolean
      mdlHelper_FolderExists = False
      On Error Resume Next
      If Not Dir(strFilePath, vbDirectory) = vbNullString And Not strFilePath = vbNullString Then
          mdlHelper_FolderExists = True
      End If
End Function

Public Function clearPath(ByVal strPath) As String
    clearPath = VBA.Replace(strPath, "\\", "\")
End Function
