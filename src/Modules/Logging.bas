Attribute VB_Name = "Logging"
Option Explicit

Public Sub Log(ByVal msg As String)
    Debug.Print Now & " - " & msg
End Sub
