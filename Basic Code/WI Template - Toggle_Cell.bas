Attribute VB_Name = "Toggle_Cell"
Public Sub ToggleCellColor(ByVal cell As Range)
    Select Case cell.Interior.Color
        Case RGB(0, 176, 80) ' green
            cell.Interior.Color = RGB(255, 0, 0) ' red
        Case RGB(255, 0, 0) ' red
            cell.Interior.Color = RGB(0, 176, 80) ' green
        Case Else
            cell.Interior.Color = RGB(0, 176, 80) ' default to green
    End Select
End Sub

