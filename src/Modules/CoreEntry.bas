Attribute VB_Name = "CoreEntry"
Option Explicit

Public Sub RunProcess(ByVal config As AppConfig)
    On Error GoTo Fail
    
    Call ProcessEngine_Run(config)
    
    Exit Sub
Fail:
    HandleError Err, "RunProcess"
End Sub
