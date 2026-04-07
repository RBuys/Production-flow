Attribute VB_Name = "IndexService"
Option Explicit

Public Function BuildIndex(ByVal rootPath As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    ' build index logic
    
    Set BuildIndex = dict
End Function
