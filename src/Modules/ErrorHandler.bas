Attribute VB_Name = "ErrorHandler"
Option Explicit

Public Sub HandleError(ByVal errObj As ErrObject, ByVal context As String)
    MsgBox "Error in " & context & ": " & errObj.Description
End Sub
