Attribute VB_Name = "modExport"
Option Explicit

Public Sub ExportAllModules()
    Dim vbComp As Object
    Dim basePath As String
    
    basePath = ThisWorkbook.path & "\src\"
    Debug.Print basePath
    EnsureFolder basePath
    EnsureFolder basePath & "Modules\"
    EnsureFolder basePath & "Classes\"
    EnsureFolder basePath & "Forms\"
    
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        
        Select Case vbComp.Type
        
            Case 1 ' Standard Module
                ExportComponent vbComp, basePath & "Modules\", ".bas"
                
            Case 2 ' Class Module
                ExportComponent vbComp, basePath & "Classes\", ".cls"
                
            Case 3 ' UserForm
                ExportComponent vbComp, basePath & "Forms\", ".frm"
                
            Case 100 ' Document (ThisWorkbook / Sheets)
                ExportComponent vbComp, basePath, ".cls"
                
        End Select
        
    Next vbComp
    
    MsgBox "? Export complete!", vbInformation
End Sub

Private Sub ExportComponent(vbComp As Object, folder As String, ext As String)
    Dim filePath As String
    
    filePath = folder & vbComp.Name & ext
    
    Debug.Print "Exporting: " & filePath
    
    If Dir(filePath) <> "" Then Kill filePath
    
    On Error GoTo ExportError
    vbComp.Export filePath
    Exit Sub

ExportError:
    MsgBox "Failed to export: " & vbComp.Name & vbCrLf & Err.Description
End Sub

Private Sub EnsureFolder(path As String)
    If Dir(path, vbDirectory) = "" Then
        On Error Resume Next
        MkDir path
        On Error GoTo 0
    End If
End Sub
