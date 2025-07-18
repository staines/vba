Attribute VB_Name = "modCopyFile_Prompt"
Option Compare Database
' created by Chris Staines
'
' required reference for complete functionality:
'   Microsoft Scripting Runtime

Public Function CopyFile_Prompt(strSource As String, strDestination As String, boolPromptIfError As Boolean) As Boolean
' attempt to copy a file and return error (prompted or not) if unable to

If Dir(strSource$) = vbNullString Then
' if the source file does not exist, ...

    If boolPromptIfError = True Then
    ' if user wishes to be prompted on error, ...
    
        MsgBox "Source file does not exist." & vbNewLine & vbNewLine & strSource$, vbCritical + vbOKOnly
        ' inform user of issue
        
    End If
    
    CopyFile_Prompt = False
    ' return error to origin function
    
Else
' if the source file exists, ...

    Dim fsObject As Object
    ' variable for file system object
    
    Set fsObject = VBA.CreateObject("Scripting.FileSystemObject")
    ' create a file system object to use

    fsObject.CopyFile strSource$, strDestination$, True
    ' attempt to copy the file
    
    If Dir(strDestination$) = vbNullString Then
    ' if the destination file does not exist, ...
    
        If boolPromptIfError = True Then
        ' if user wishes to be prompted on error, ...
        
            MsgBox "Destination file does not exist after attempted copy." & vbNewLine & vbNewLine & strDestination$, vbCritical + vbOKOnly
            ' inform user of issue
            
        End If
        
        CopyFile_Prompt = False
        ' return error to origin function
        
    Else
    ' if the destination file does exist, ...
    
        CopyFile_Prompt = True
        ' return success to origin function
        
    End If

End If

End Function
