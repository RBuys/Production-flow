Attribute VB_Name = "FileSystemService"
Option Explicit

Public Function GetFSO() As Object
    Set GetFSO = CreateObject("Scripting.FileSystemObject")
End Function
