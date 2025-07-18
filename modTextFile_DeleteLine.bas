Attribute VB_Name = "modTextFile_DeleteLine"
'
' created by Chris Staines
'
' deletes a number of lines from the beginning of a text file
' useful when normalizing outdated export filetypes

Public Function TextFile_DeleteLine(strFile As String, lngLineToDelete As Long, _
    Optional strCheckFirstCharactersFor As String, Optional boolSkipLastLine As Boolean)
' delete the given amount of lines from the beginning of a text file

Dim lngLineIndex As Long
' variable for line index in file when writing

If Dir(strFile$) <> vbNullString Then
' if the file exists, ...
    
    Dim fsoFileSystem As FileSystemObject
    ' variable for file system object
    
    Dim tsFile As TextStream
    ' variable for file
    
    Dim arrTextFile As Variant
    
    Set fsoFileSystem = CreateObject("Scripting.FileSystemObject")
    ' open file system
    
    Set tsFile = fsoFileSystem.OpenTextFile(strFile$, ForReading, TristateFalse)
    ' open text stream for reading
    
    If tsFile.AtEndOfStream = False Then
    ' if the file is not blank, ...
        
        arrTextFile = Split(tsFile.ReadAll, vbNewLine)
        ' obtain text file data, and split by new line
        
        tsFile.Close
        ' close the text stream

        If strCheckFirstCharactersFor$ = vbNullString Or _
            Left(arrTextFile(0), Len(strCheckFirstCharactersFor$)) = strCheckFirstCharactersFor$ Then
        ' if there is no check characters, or if the check characters match, then...

            If UBound(arrTextFile) >= (lngLineToDelete& - 1) Then
            ' if there are at least as many lines in the file as lines to be deleted, ...
                
                Set tsFile = fsoFileSystem.OpenTextFile(strFile$, ForWriting, False, TristateFalse)
                ' open text stream for writing
    
                If boolSkipLastLine = True Then lngLineCount& = UBound(arrTextFile) - 1 Else lngLineCount& = UBound(arrTextFile)
                ' if the user requests to skip the last line of the file, then do so; otherwise, do not
                
                For lngLineIndex& = (lngLineToDelete&) To lngLineCount&
                ' from the line after the line to delete, to the end of the lines in the array, ...
                
                    tsFile.WriteLine arrTextFile(lngLineIndex&)
                    ' write the line
                
                Next lngLineIndex&
                ' continue to the next line
                
            End If

        End If

    End If
    
    tsFile.Close
    ' close the text stream
    
    Set tsFile = Nothing
    ' clear from memory
    
    Set fsoFileSystem = Nothing
    ' clear from memory

End If

End Function

