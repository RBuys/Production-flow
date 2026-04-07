Attribute VB_Name = "FolderRules"
Option Explicit

Public Function IsExcludedFolder(ByVal folderName As String) As Boolean
    Select Case LCase(folderName)
        Case "archive", "old", "temp"
            IsExcludedFolder = True
    End Select
End Function
