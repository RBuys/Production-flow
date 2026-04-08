Attribute VB_Name = "modImport"
Option Explicit

Public Sub ImportAllModules()

    Dim basePath As String
    
    basePath = ThisWorkbook.path & "\src\"
    
    Debug.Print "Base Path: " & basePath
    
    If Dir(basePath, vbDirectory) = "" Then
        MsgBox "Base folder not found: " & basePath, vbCritical
        Exit Sub
    End If
    
    If Not HasVBFiles(basePath) Then
        MsgBox "No VBA source files found. Import cancelled.", vbCritical
        Exit Sub
    End If
    
    If MsgBox("This will replace all VBA modules. Continue?", vbYesNo + vbExclamation) = vbNo Then
        Exit Sub
    End If
    
    RemoveAllModules
    
    ImportFolder basePath & "Modules\"
    ImportFolder basePath & "Classes\"
    ImportFolder basePath & "Forms\"
    
    MsgBox "Import complete.", vbInformation

End Sub


Private Sub ImportFolder(folderPath As String)

    Dim file As String
    Dim currentModuleName As String
    
    currentModuleName = "modImport"
    
    If Dir(folderPath, vbDirectory) = "" Then
        Debug.Print "Folder missing: " & folderPath
        Exit Sub
    End If
    
    Debug.Print "Importing from: " & folderPath
    
    file = Dir(folderPath & "*.*")
    
    Do While file <> ""
    
        If IsValidVBFile(file) Then
        
            ' Skip importing the module that is running this code
            If LCase(file) = LCase(currentModuleName & ".bas") Then
                Debug.Print "Skipping self import: " & file
                GoTo NextFile
            End If
        
            Debug.Print "Importing: " & folderPath & file
            
            On Error GoTo ImportError
            
            ThisWorkbook.VBProject.VBComponents.Import folderPath & file
            
            On Error GoTo 0
            
        End If
        
NextFile:
        file = Dir
        
    Loop
    
    Exit Sub

ImportError:

    MsgBox "Failed to import: " & file & vbCrLf & Err.Description, vbCritical
    Resume Next

End Sub


Private Function IsValidVBFile(fileName As String) As Boolean

    Dim ext As String
    
    ext = LCase(Right(fileName, 4))
    
    Select Case ext
    
        Case ".bas", ".cls", ".frm"
            IsValidVBFile = True
            
        Case Else
            IsValidVBFile = False
            
    End Select

End Function


Private Function HasVBFiles(basePath As String) As Boolean

    If Dir(basePath & "Modules\*.bas") <> "" Then
        HasVBFiles = True
        Exit Function
    End If
    
    If Dir(basePath & "Classes\*.cls") <> "" Then
        HasVBFiles = True
        Exit Function
    End If
    
    If Dir(basePath & "Forms\*.frm") <> "" Then
        HasVBFiles = True
        Exit Function
    End If
    
    HasVBFiles = False

End Function


Private Sub RemoveAllModules()

    Dim vbComp As Object
    Dim i As Long
    Dim thisModuleName As String
    
    thisModuleName = "modImport"
    
    Debug.Print "Removing modules (except " & thisModuleName & ")"
    
    For i = ThisWorkbook.VBProject.VBComponents.Count To 1 Step -1
    
        Set vbComp = ThisWorkbook.VBProject.VBComponents(i)
        
        If vbComp.Name = thisModuleName Then
            Debug.Print "Skipping self: " & vbComp.Name
            GoTo NextComponent
        End If
        
        If vbComp.Type = 1 Or vbComp.Type = 2 Or vbComp.Type = 3 Then
        
            Debug.Print "Removing: " & vbComp.Name
            
            On Error Resume Next
            ThisWorkbook.VBProject.VBComponents.Remove vbComp
            On Error GoTo 0
            
        End If
        
NextComponent:
    
    Next i

End Sub
