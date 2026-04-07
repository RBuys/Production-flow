Attribute VB_Name = "PushDate"
Option Explicit

' ============================================================
' ENTRY POINT
' ============================================================
Private Const DEBUG_KEEP_JC_OPEN As Boolean = False

Public Sub PushOrdersBackToJC()
    Dim reportSheet As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim jobNumber As String
    Dim materialText As String
    Dim requiredDate As Variant
    Dim jcPath As String
    Dim updatedCount As Long
    Dim ordNum As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set reportSheet = ThisWorkbook.Worksheets(1)
    lastRow = reportSheet.Cells(reportSheet.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        jobNumber = Trim$(CStr(reportSheet.Cells(r, "A").Value))
        materialText = Trim$(CStr(reportSheet.Cells(r, "B").Value))
        requiredDate = reportSheet.Cells(r, "D").Value
        ordNum = Trim$(CStr(reportSheet.Cells(r, "C").Value))
        
        If jobNumber = "" Or materialText = "" Then GoTo NextRow
        If IsEmpty(requiredDate) Or requiredDate = "" Then GoTo NextRow

        jcPath = FindJCFile(jobNumber)
        If jcPath = "" Then
            SavePendingPush jobNumber, materialText, ordNum, _
                            requiredDate, "JC file not found"
            GoTo NextRow
        End If

        If WriteRequiredDateToJC(jcPath, materialText, requiredDate) Then
            updatedCount = updatedCount + 1
            ClearPendingPush jobNumber, materialText
        Else
            SavePendingPush jobNumber, materialText, ordNum, _
                            requiredDate, "JC file locked or read-only"
        End If

NextRow:
    Next r

CleanExit:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox updatedCount & " received date(s) sent to JC files.", _
           vbInformation, "Push Back Complete"
    Exit Sub

CleanFail:
    MsgBox Err.Description, vbCritical, "Push Back Failed"
    Resume CleanExit
End Sub

' ============================================================
' FILE LOOKUP
' ============================================================
Private Function FindJCFile(ByVal jobNumber As String) As String
    FindJCFile = FindExcelFileRecursive(ResolveWorkshopPath(), jobNumber)
End Function

Private Function FindExcelFileRecursive(ByVal folderPath As String, _
                                        ByVal jobNumber As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object
    Dim ext As String

    Set fso = GetFSO()
    If Not fso.FolderExists(folderPath) Then Exit Function
    Set folder = fso.GetFolder(folderPath)

    If IsExcludedFolder(folder.Name) Then Exit Function

    For Each file In folder.Files
        If fso.GetBaseName(file.Name) = jobNumber Then
            ext = LCase$(fso.GetExtensionName(file.Name))
            Select Case ext
                Case "xlsx", "xlsm", "xls", "xlsb"
                    FindExcelFileRecursive = file.Path
                    Exit Function
            End Select
        End If
    Next file

    For Each subFolder In folder.SubFolders
        FindExcelFileRecursive = FindExcelFileRecursive(subFolder.Path, jobNumber)
        If FindExcelFileRecursive <> "" Then Exit Function
    Next subFolder
End Function


Private Function FindFileRecursive(ByVal folderPath As String, _
                                   ByVal fileName As String) As String
    Dim fso As Object
    Dim folder As Object
    Dim subFolder As Object
    Dim file As Object

    Set fso = GetFSO()
    If Not fso.FolderExists(folderPath) Then Exit Function
    Set folder = fso.GetFolder(folderPath)

    ' Skip excluded folders entirely
    If IsExcludedFolder(folder.Name) Then Exit Function

    For Each file In folder.Files
        If LCase$(file.Name) = LCase$(fileName) Then
            FindFileRecursive = file.Path
            Exit Function
        End If
    Next file

    For Each subFolder In folder.SubFolders
        FindFileRecursive = FindFileRecursive(subFolder.Path, fileName)
        If FindFileRecursive <> "" Then Exit Function
    Next subFolder
End Function


' ============================================================
' JC UPDATE — REQUIRED DATE ONLY (COLUMN P)
' ============================================================
Private Function WriteRequiredDateToJC(ByVal jcPath As String, _
                                       ByVal materialText As String, _
                                       ByVal requiredDate As Variant) As Boolean
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim r As Long
    Dim jcMaterial As String

    On Error Resume Next
    Set wb = Workbooks.Open(jcPath, ReadOnly:=False, UpdateLinks:=False)
    If wb Is Nothing Then Exit Function
    On Error GoTo 0

    If wb.ReadOnly Then
        If Not DEBUG_KEEP_JC_OPEN Then wb.Close SaveChanges:=False
        WriteRequiredDateToJC = False
        Exit Function
    End If

    For Each ws In wb.Worksheets
        For r = 9 To 38
            jcMaterial = Trim$(CStr(ws.Cells(r, "J").Value))

            If jcMaterial = materialText Then

                ws.Cells(r, "P").Value = requiredDate
                wb.Save

                If Not DEBUG_KEEP_JC_OPEN Then
                    wb.Close SaveChanges:=True
                End If

                WriteRequiredDateToJC = True
                Exit Function
            End If
        Next r
    Next ws

    If Not DEBUG_KEEP_JC_OPEN Then
        wb.Close SaveChanges:=False
    End If

    WriteRequiredDateToJC = False
End Function

' ============================================================
' PENDING PUSH STORAGE
' ============================================================
Private Sub SavePendingPush(ByVal jobNum As String, _
                            ByVal material As String, _
                            ByVal orderNum As String, _
                            ByVal reqDate As Variant, _
                            ByVal reason As String)
    Dim ws As Worksheet
    Dim r As Long
    Dim lastRow As Long

    Set ws = GetPendingSheet()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For r = 2 To lastRow
        If ws.Cells(r, 1).Value = jobNum _
           And ws.Cells(r, 2).Value = material Then

            ws.Cells(r, 3).Value = orderNum
            ws.Cells(r, 4).Value = reqDate
            ws.Cells(r, 5).Value = Now
            ws.Cells(r, 6).Value = reason
            Exit Sub
        End If
    Next r

    ws.Cells(lastRow + 1, 1).Resize(1, 6).Value = _
        Array(jobNum, material, orderNum, reqDate, Now, reason)
End Sub

Private Sub ClearPendingPush(ByVal jobNum As String, ByVal material As String)
    Dim ws As Worksheet
    Dim r As Long

    Set ws = GetPendingSheet()

    For r = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row To 2 Step -1
        If ws.Cells(r, 1).Value = jobNum _
           And ws.Cells(r, 2).Value = material Then
            ws.Rows(r).Delete
            Exit Sub
        End If
    Next r
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
' SHARED FILESYSTEMOBJECT
' ============================================================
Private Function GetFSO() As Object
    Static fso As Object
    If fso Is Nothing Then Set fso = CreateObject("Scripting.FileSystemObject")
    Set GetFSO = fso
End Function



