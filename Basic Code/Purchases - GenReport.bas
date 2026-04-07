Attribute VB_Name = "GenReport"
Option Explicit

Public Sub GenerateReport()
    Dim targetSheet As Worksheet
    Dim pendingSheet As Worksheet
    Dim pendingDict As Object
    Dim writeRow As Long
    Dim lastRow As Long
    Dim workshopPath As String
    Dim jcPath As Variant
    Dim r As Long
    Dim key As String

    On Error GoTo CleanFail

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    Set targetSheet = ThisWorkbook.Worksheets(1)
    Set pendingSheet = GetPendingSheet()
    Set pendingDict = CreateObject("Scripting.Dictionary")

    ' Load pending pushes
    lastRow = pendingSheet.Cells(pendingSheet.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        key = pendingSheet.Cells(r, "A").Value & "|" & pendingSheet.Cells(r, "B").Value
        If Not pendingDict.Exists(key) Then
            pendingDict.Add key, Array( _
                pendingSheet.Cells(r, "C").Value, _
                pendingSheet.Cells(r, "D").Value)
        End If
    Next r

    ' Clear report
    targetSheet.Range("A2:D" & targetSheet.Rows.Count).ClearContents
    writeRow = 2

    ' Build JC index ONCE
    workshopPath = ResolveWorkshopPath()
    BuildJCIndex workshopPath

    ' Process all JCs
    For Each jcPath In GetAllJCPaths()
        ProcessJCWorkbook jcPath, targetSheet, writeRow
    Next jcPath

    ' Re-inject pending values
    lastRow = targetSheet.Cells(targetSheet.Rows.Count, "A").End(xlUp).Row
    For r = 2 To lastRow
        key = targetSheet.Cells(r, "A").Value & "|" & targetSheet.Cells(r, "B").Value
        If pendingDict.Exists(key) Then
            targetSheet.Cells(r, "C").Value = pendingDict(key)(0)
            targetSheet.Cells(r, "D").Value = pendingDict(key)(1)
        End If
    Next r

    ' Sort
    If lastRow >= 2 Then
        With targetSheet.Sort
            .SortFields.Clear
            .SortFields.Add targetSheet.Range("A2:A" & lastRow), xlSortOnValues, xlAscending
            .SetRange targetSheet.Range("A1:D" & lastRow)
            .Header = xlYes
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
    MsgBox Err.Description, vbCritical
    Resume CleanExit
End Sub


