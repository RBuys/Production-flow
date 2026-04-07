Attribute VB_Name = "JCIndex"
Option Explicit

Private JCIndex As Object

Public Sub BuildJCIndex(ByVal workshopPath As String)
    Set JCIndex = CreateObject("Scripting.Dictionary")
    ScanFolderForIndex workshopPath
End Sub

Public Function GetJCPath(ByVal jobNumber As String) As String
    If JCIndex.Exists(jobNumber) Then
        GetJCPath = JCIndex(jobNumber)
    Else
        GetJCPath = ""
    End If
End Function

Private Sub ScanFolderForIndex(ByVal folderPath As String)
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim jobNum As String
    Dim ext As String

    Set fso = GetFSO()
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set folder = fso.GetFolder(folderPath)

    If IsExcludedFolder(folder.Name) Then Exit Sub

    For Each file In folder.Files
        ext = LCase$(fso.GetExtensionName(file.Name))
        Select Case ext
            Case "xlsx", "xlsm", "xls", "xlsb"
                jobNum = fso.GetBaseName(file.Name)
                If Not JCIndex.Exists(jobNum) Then
                    JCIndex.Add jobNum, file.Path
                End If
        End Select
    Next file

    For Each subFolder In folder.SubFolders
        ScanFolderForIndex subFolder.Path
    Next subFolder
End Sub

