Attribute VB_Name = "PushBack"
Option Explicit
Private Const DEBUG_KEEP_JC_OPEN As Boolean = False

Public Sub PushOrdersBackToJC()
    Dim reportSheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim jobNumber As String
    Dim materialText As String
    Dim orderNumber As String
    Dim requiredDate As Variant
    Dim jcPath As String
    Dim updatedCount As Long

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set reportSheet = ThisWorkbook.Worksheets(1)
    lastRow = reportSheet.Cells(reportSheet.Rows.Count, "A").End(xlUp).Row

    ' Build JC index ONCE
    BuildJCIndex ResolveWorkshopPath()

    For r = 2 To lastRow
        jobNumber = Trim$(CStr(reportSheet.Cells(r, "A").Value))
        materialText = Trim$(CStr(reportSheet.Cells(r, "B").Value))
        orderNumber = Trim$(CStr(reportSheet.Cells(r, "C").Value))
        requiredDate = reportSheet.Cells(r, "D").Value

        If jobNumber = "" Or materialText = "" Or orderNumber = "" Then GoTo NextRow

        jcPath = GetJCPath(jobNumber)
        If jcPath = "" Then
            SavePendingPush jobNumber, materialText, orderNumber, requiredDate, "JC file not found"
            GoTo NextRow
        End If

        If WriteOrderToJC(jcPath, materialText, orderNumber, requiredDate) Then
            updatedCount = updatedCount + 1
            ClearPendingPush jobNumber, materialText
        Else
            SavePendingPush jobNumber, materialText, orderNumber, requiredDate, "JC file locked or read-only"
        End If
NextRow:
    Next r

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox updatedCount & " order(s) pushed back.", vbInformation
    Exit Sub

CleanFail:
    MsgBox Err.Description, vbCritical
    Resume CleanExit
End Sub

