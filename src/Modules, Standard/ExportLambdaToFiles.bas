Sub ExportLambdasToFiles()
    Dim nme As Name
    Dim FolderPath As String
    Dim FolderPicker As FileDialog
    Dim FormulaText As String
    Dim CommentText As String
    Dim FilePath As String
    Dim FileNum As Integer
    Dim ExportCount As Long
    Dim LambdaPos As Long
    
    ' 1. Prompt user to select output folder
    Set FolderPicker = Application.FileDialog(msoFileDialogFolderPicker)
    With FolderPicker
        .Title = "Select Folder to Export Lambda Functions"
        .AllowMultiSelect = False
        If .Show = -1 Then
            FolderPath = .SelectedItems(1)
        Else
            MsgBox "Export canceled.", vbInformation
            Exit Sub
        End If
    End With
    
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    ExportCount = 0
    
    ' 2. Iterate through all Defined Names
    For Each nme In ThisWorkbook.Names
        FormulaText = nme.RefersTo
        LambdaPos = InStr(1, FormulaText, "LAMBDA(", vbTextCompare)
        
        ' Check if it is a valid LAMBDA definition
        If LambdaPos > 0 Then
            
            ' Extract Name Manager comment
            CommentText = ""
            On Error Resume Next
            CommentText = nme.Comment
            On Error GoTo 0
            
            ' If a comment exists, inject it right after "LAMBDA("
            If Trim(CommentText) <> "" Then
                ' Escape double quotes inside the comment string
                CommentText = Replace(CommentText, """", """""")
                
                ' Construct the parameter-level comment assignment
                ' Resulting insertion: _comment, N("the comment"),
                FormulaText = Left(FormulaText, LambdaPos + 6) & vbCrLf & _
                              "    _comment, N(""" & CommentText & """)," & vbCrLf & _
                              Mid(FormulaText, LambdaPos + 7)
            End If
            
            ' 3. Save to .xlsxfx file
            FilePath = FolderPath & nme.Name & ".xlsxfx"
            FileNum = FreeFile
            Open FilePath For Output As #FileNum
            Print #FileNum, FormulaText
            Close #FileNum
            
            ExportCount = ExportCount + 1
        End If
    Next nme
    
    MsgBox ExportCount & " Lambda function(s) successfully exported to:" & vbCrLf & FolderPath, vbInformation
End Sub