Attribute VB_Name = "Module1"
Public Function IsExcludedFolder(ByVal folderName As String) As Boolean
    Select Case LCase$(folderName)
        Case "1 - ncr", _
             "2 - rework", _
             "6 - dispatch", _
             "99 - templates"
            IsExcludedFolder = True
    End Select
End Function


