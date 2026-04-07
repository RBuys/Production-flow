Attribute VB_Name = "ProcessEngine"
Option Explicit

Public Sub ProcessEngine_Run(ByVal config As AppConfig)
    Dim ws As Worksheet
    Dim pendingWs As Worksheet
    
    Set ws = ThisWorkbook.Worksheets(config.TargetSheetName)
    Set pendingWs = ThisWorkbook.Worksheets(config.PendingSheetName)
    
    Select Case config.Mode
        Case "PURCHASES"
            ProcessPurchases ws, pendingWs, config
        Case "RECEIVABLES"
            ProcessReceivables ws, pendingWs, config
        Case "OVERDUE"
            ProcessOverdue ws, pendingWs, config
        Case Else
            Err.Raise vbObjectError + 1, , "Invalid mode"
    End Select
End Sub
