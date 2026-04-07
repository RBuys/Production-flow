Attribute VB_Name = "modFolderRules"
Public Function IsExcludedFolder(ByVal folderName As String) As Boolean
    Select Case LCase$(folderName)
        Case "6 - dispatch", _
             "99 - templates", _
             "1 - ncr", _
             "2 - rework"
            IsExcludedFolder = True
    End Select
End Function


