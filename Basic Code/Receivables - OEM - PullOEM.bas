Attribute VB_Name = "PullOEM"
Option Explicit

' ============================================================
' ENTRY POINT – GENERATE REPORT
' ============================================================
Public Sub GenerateReport()
    Dim targetSheet As Worksheet
    Dim pendingSheet As Worksheet
    Dim writeRow As Long
    Dim lastRow As Long
    Dim workshopPath As String
    Dim pendingDict As Object
    Dim r As Long
    Dim key As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set targetSheet = ThisWorkbook.Worksheets(1)
    Set pendingSheet = GetPendingSheet()
    Set pendingDict = CreateObject("Scripting.Dictionary")

    ' --------------------------------------------------------
    ' 1. Load pending pushes into memory
    ' --------------------------------------------------------
    lastRow = pendingSheet.Cells(pendingSheet.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        key = pendingSheet.Cells(r, "A").Value & "|" & pendingSheet.Cells(r, "B").Value
        If Not pendingDict.Exists(key) Then
            pendingDict.Add key, Array( _
                pendingSheet.Cells(r, "C").Value, _
                pendingSheet.Cells(r, "D").Value)
        End If
    Next r

    ' --------------------------------------------------------
    ' 2. Clear visible report
    ' --------------------------------------------------------
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    If lastRow >= 2 Then
        targetSheet.Range("A2:D" & lastRow).ClearContents
    End If
    writeRow = 2

    ' --------------------------------------------------------
    ' 3. Resolve Workshop path
    ' --------------------------------------------------------
    workshopPath = ResolveWorkshopPath()

    ' --------------------------------------------------------
    ' 4. Scan JC files
    ' --------------------------------------------------------
    ScanFolder workshopPath, targetSheet, writeRow

    ' --------------------------------------------------------
    ' 5. Re-inject pending order numbers
    ' --------------------------------------------------------
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        key = targetSheet.Cells(r, "A").Value & "|" & targetSheet.Cells(r, "B").Value
        If pendingDict.Exists(key) Then
            targetSheet.Cells(r, "C").Value = pendingDict(key)(0)
            targetSheet.Cells(r, "D").Value = pendingDict(key)(1)
        End If
    Next r

    ' --------------------------------------------------------
    ' 6. Sort report by Job # (Column A)
    ' --------------------------------------------------------
    If lastRow >= 2 Then
        With targetSheet.Sort
            .SortFields.Clear
            .SortFields.Add key:=targetSheet.Range("A2:A" & lastRow), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .SetRange targetSheet.Range("A1:D" & lastRow)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Report generated successfully.", vbInformation
    Exit Sub

CleanFail:
    MsgBox Err.Description, vbCritical, "Report Failed"
    Resume CleanExit
End Sub

' ============================================================
' SAFE RECURSIVE FOLDER SCAN
' ============================================================
Private Sub ScanFolder(ByVal folderPath As String, _
                       ByVal targetSheet As Worksheet, _
                       ByRef writeRow As Long)

    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object

    Set fso = GetFSO()
    If Not fso.FolderExists(folderPath) Then Exit Sub
    Set folder = fso.GetFolder(folderPath)

    ' Skip excluded folders entirely
    If IsExcludedFolder(folder.Name) Then Exit Sub

    ' Process ALL .xlsx files
    For Each file In folder.Files
    Select Case LCase$(fso.GetExtensionName(file.Name))
        Case "xlsx", "xlsm", "xls", "xlsb"
            ProcessJCWorkbook file.Path, targetSheet, writeRow
    End Select
    Next file

    ' Recurse
    For Each subFolder In folder.SubFolders
        ScanFolder subFolder.Path, targetSheet, writeRow
    Next subFolder
End Sub


' ============================================================
' JC WORKBOOK PROCESSING (READ-ONLY)
' ============================================================
Private Sub ProcessJCWorkbook(ByVal filePath As String, _
                              ByVal targetSheet As Worksheet, _
                              ByRef writeRow As Long)

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim data As Variant
    Dim r As Long
    Dim jobNumber As String
    Dim materialText As String
    Dim normalizedText As String
    Dim reportValueB As String

    jobNumber = GetFileBaseName(filePath)

    On Error Resume Next
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    If wb Is Nothing Then Exit Sub
    On Error GoTo 0

    For Each ws In wb.Worksheets

        data = ws.Range("E9:L38").Value   ' E=1, J=6, L=8

        For r = 1 To UBound(data, 1)

            materialText = Trim$(CStr(data(r, 6))) ' JC column J (OEM check)
            If materialText = "" Then GoTo NextRow
            If Trim$(CStr(data(r, 8))) <> "" Then GoTo NextRow

            normalizedText = Replace(UCase$(materialText), ".", "")
            normalizedText = Replace(normalizedText, " ", "")
            If InStr(normalizedText, "OEM") = 0 Then GoTo NextRow

            reportValueB = Trim$(CStr(data(r, 1))) ' JC column E

            targetSheet.Cells(writeRow, 1).Value = jobNumber
            targetSheet.Cells(writeRow, 2).Value = reportValueB
            writeRow = writeRow + 1

NextRow:
        Next r
    Next ws

    wb.Close SaveChanges:=False
End Sub

' ============================================================
' WORKSHOP PATH RESOLUTION
' ============================================================
Public Function ResolveWorkshopPath() As String
    Dim fso As Object
    Dim currentPath As String
    Dim parentPath As String

    Set fso = GetFSO()
    currentPath = ThisWorkbook.Path

    Do While currentPath <> ""
        If LCase$(fso.GetFileName(currentPath)) = "workshop" Then
            ResolveWorkshopPath = currentPath
            Exit Function
        End If
        parentPath = fso.GetParentFolderName(currentPath)
        If parentPath = currentPath Then Exit Do
        currentPath = parentPath
    Loop

    Err.Raise vbObjectError + 1000, _
              "ResolveWorkshopPath", _
              "Workshop folder could not be found from: " & ThisWorkbook.Path
End Function

' ============================================================
' HIDDEN PENDING PUSH SHEET
' ============================================================
Public Function GetPendingSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_PendingPush")
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        With ws
            .Name = "_PendingPush"
            .Range("A1:F1").Value = Array( _
                "JobNumber", "Material", "OrderNumber", _
                "RequiredDate", "LastAttempt", "FailureReason")
            .Columns("A:F").NumberFormat = "@"
            .Columns("E").NumberFormat = "yyyy-mm-dd hh:mm:ss"
            .Visible = xlSheetVeryHidden
        End With
    End If

    Set GetPendingSheet = ws
End Function

' ============================================================
' UTILITIES
' ============================================================
Private Function GetFileBaseName(ByVal fullPath As String) As String
    GetFileBaseName = GetFSO().GetBaseName(fullPath)
End Function

Private Function GetFSO() As Object
    Static fso As Object
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = fso
End Function



