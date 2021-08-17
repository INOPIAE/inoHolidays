Attribute VB_Name = "mdl_Translation"
Option Explicit

Public strLabel(7) As String
Public strScreentip(5) As String
Public strSupertip(5) As String
Public strError(5) As String
Public strfrmInfo(0) As String

Public Sub germanText()
    strfrmInfo(0) = "Der Quellcode is OpenSource unter AGPLv3 und verfügbar auf "
End Sub

Public Sub englishText()
    strfrmInfo(0) = "Source code is OpenSource under AGPLv3 and available at "
End Sub

