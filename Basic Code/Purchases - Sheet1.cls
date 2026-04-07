VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim changedCell As Range

    ' Only react to changes in Column C (Order #), rows 2 down
    If Intersect(Target, Me.Columns("C")) Is Nothing Then Exit Sub
    If Target.Row < 2 Then Exit Sub
    If Target.Cells.Count > 1 Then Exit Sub

    Application.EnableEvents = False

    Set changedCell = Target

    ' If Order # entered ? set date 3 working days later
    If Trim(CStr(changedCell.Value)) <> "" Then
        changedCell.Offset(0, 1).Value = WorksheetFunction.WorkDay(Date, 3)
    Else
        ' If Order # cleared ? clear date
        changedCell.Offset(0, 1).ClearContents
    End If

    Application.EnableEvents = True
End Sub

