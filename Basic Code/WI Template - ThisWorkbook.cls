VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
    If Target.Cells.CountLarge <> 1 Then Exit Sub
    If Intersect(Target, Sh.Range("G9:G43")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error GoTo CleanExit

    Call ToggleCellColor(Target)

CleanExit:
    Application.EnableEvents = True
End Sub

