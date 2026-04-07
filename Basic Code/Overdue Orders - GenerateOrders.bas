Attribute VB_Name = "GenerateOrders"
Option Explicit

' ============================================================
' ENTRY POINT – GENERATE OVERDUE ORDERS REPORT (EXACT DAYS)
' ============================================================
Public Sub GenerateOverdueOrdersReport()
    Dim targetSheet As Worksheet
    Dim writeRow As Long
    Dim workshopPath As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set targetSheet = ThisWorkbook.Worksheets(1)

    targetSheet.Range("A2:E" & targetSheet.Rows.Count).ClearContents
    writeRow = 2

    workshopPath = ResolveWorkshopPath()
    ScanFolder workshopPath, targetSheet, writeRow

   ' Sort by Days Overdue (desc), then Job Number (asc)
    If writeRow > 2 Then
        With targetSheet.Sort
            .SortFields.Clear
    
            ' Primary: Days Overdue (highest first)
            .SortFields.Add _
                Key:=targetSheet.Range("E2:E" & writeRow - 1), _
                SortOn:=xlSortOnValues, _
                Order:=xlDescending
    
            ' Secondary: Job Number (A–Z)
            .SortFields.Add _
                Key:=targetSheet.Range("A2:A" & writeRow - 1), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending
    
            .SetRange targetSheet.Range("A1:E" & writeRow - 1)
            .Header = xlYes
            .Apply
        End With
    End If

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    MsgBox "Overdue orders report generated.", vbInformation
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

    ' Process all .xlsx files
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
    Dim material As String
    Dim orderNum As String
    Dim requiredDate As Date
    Dim receivedDate As Variant
    Dim daysOverdue As Long
    Dim normalized As String

    jobNumber = GetFileBaseName(filePath)

    On Error Resume Next
    Set wb = Workbooks.Open(filePath, ReadOnly:=True, UpdateLinks:=False, AddToMru:=False)
    If wb Is Nothing Then Exit Sub
    On Error GoTo 0

    For Each ws In wb.Worksheets
        ' J=1, L=3, P=7, Q=8
        data = ws.Range("J9:Q38").Value

        For r = 1 To UBound(data, 1)

            material = Trim$(CStr(data(r, 1)))
            orderNum = Trim$(CStr(data(r, 3)))
            receivedDate = data(r, 7)

            If material = "" Then GoTo NextRow
            If orderNum = "" Then GoTo NextRow
            If Not IsEmpty(receivedDate) And receivedDate <> "" Then GoTo NextRow
            If Not IsDate(data(r, 8)) Then GoTo NextRow

            requiredDate = CDate(data(r, 8))
            If requiredDate > Date Then GoTo NextRow

            normalized = Replace(UCase$(material), ".", "")
            normalized = Replace(normalized, " ", "")
            If InStr(normalized, "OEM") > 0 Then GoTo NextRow

            daysOverdue = Date - requiredDate

            targetSheet.Cells(writeRow, 1).Value = jobNumber
            targetSheet.Cells(writeRow, 2).Value = material
            targetSheet.Cells(writeRow, 3).Value = orderNum
            targetSheet.Cells(writeRow, 4).Value = requiredDate
            targetSheet.Cells(writeRow, 5).Value = daysOverdue
            writeRow = writeRow + 1

NextRow:
        Next r
    Next ws

    wb.Close SaveChanges:=False
End Sub

' ============================================================
' PATH RESOLUTION
' ============================================================
Private Function ResolveWorkshopPath() As String
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


