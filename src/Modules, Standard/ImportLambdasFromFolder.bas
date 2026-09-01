Sub ImportLambdasFromFolder()
    Dim FSO As Object, Folder As Object, File As Object
    Dim FileStream As Object
    Dim FolderPath As String
    Dim FuncName As String, FuncFormula As String
    
    ' 1. Set your folder path here (must end with a backslash \)
    FolderPath = "C:\MyLambdas\" 
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    If Not FSO.FolderExists(FolderPath) Then
        MsgBox "Folder not found: " & FolderPath, vbCritical
        Exit Sub
    End If
    
    Set Folder = FSO.GetFolder(FolderPath)
    
    Application.ScreenUpdating = False
    
    For Each File In Folder.Files
        ' Process only .txt files
        If LCase(FSO.GetExtensionName(File.Path)) = "txt" Then
            
            ' Extract function name from file name (without .txt extension)
            FuncName = FSO.GetBaseName(File.Path)
            
            ' Read full formula content from the text file
            Set FileStream = FSO.OpenTextFile(File.Path, 1) ' 1 = ForReading
            FuncFormula = Trim(FileStream.ReadAll)
            FileStream.Close
            
            ' Ensure the formula starts with an '=' sign for Excel Name Manager
            If Left(FuncFormula, 1) <> "=" Then
                FuncFormula = "=" & FuncFormula
            End If
            
            ' Add or overwrite the defined name in the active workbook
            On Error Resume Next
            ThisWorkbook.Names.Add Name:=FuncName, RefersTo:=FuncFormula
            
            If Err.Number <> 0 Then
                Debug.Print "Failed to import: " & FuncName & " - " & Err.Description
                Err.Clear
            Else
                Debug.Print "Successfully imported: " & FuncName
            End If
            On Error GoTo 0
            
        End If
    Next File
    
    Application.ScreenUpdating = True
    
    MsgBox "Lambda import process complete!", vbInformation
End Sub