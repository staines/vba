Attribute VB_Name = "modReadTextFile"
' updated 2015.04.02
' created by Chris Staines
'
' required reference for complete functionality:
'   Microsoft Scripting Runtime

Public Function ReadTextFile(strFile As String) As String
' read a file and return as string

If Dir(strFile$) <> vbNullString Then
' if the file exists, ...
    
    Dim fsoFileSystem As FileSystemObject
    ' variable for file system object
    
    Dim tsFile As TextStream
    ' variable for file
    
    Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
    ' open file system
    
    Set tsFile = fsoFileSystem.OpenTextFile(strFile$, 1, TristateFalse)
    ' open catalog.out as text stream
    
    ReadTextFile = tsFile.ReadAll
    ' return file contents to origin function

    tsFile.Close
    ' close the text stream (file)
    
    Set tsFile = Nothing
    ' clear from memory
    
    Set fsoFileSystem = Nothing
    ' clear from memory

End If

End Function

