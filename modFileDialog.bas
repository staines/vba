Attribute VB_Name = "modFileDialog"
' updated 2013.09.10
' created by Chris Staines
' based on http://msdn.microsoft.com/en-us/library/ff836226.aspx
'
' open a file dialog, using options as requested by the origin function

Option Compare Database

Public Enum FileTypeIndex
' enumerator for file type index

    typAll = 1
    typExcel = 2
    typAccess = 3
    typText = 4

End Enum

Public Enum FileDialogType
' enumerator for file dialog type

    typFileDialogOpen = 1
    typFileDialogSaveAs = 2
    typFileDialogFilePicker = 3
    typFileDialogFolderPicker = 4
    
End Enum

Function FileDialog(typFileTypeIndex As FileTypeIndex, typFileDialog As FileDialogType, _
    Optional strInitialFilePath As String, Optional boolAllowMultiSelect As Boolean, Optional strTitle As String) As Variant
' open a file dialog, using options as provided by the origin function

Dim arrResult(1)
' variable for result

With Application.FileDialog(typFileDialog)
' working with a file dialog, ...

    If typFileDialog = typFileDialogFilePicker Or typFileDialog = typFileDialogOpen Then
    ' if a file picker or open dialog, ...
    
        .AllowMultiSelect = boolAllowMultiSelect
        ' allow/disallow multiselect as requested by origin function
        
        .Filters.Clear
        ' clear existing filters
        
        .Filters.Add "All Files", "*.*"
        ' all files
        
        .Filters.Add "Excel Workbooks", "*.xl*;*.csv"
        ' excel workbooks
        ' yes, *.xl*... even microsoft uses this technique in excel 2010
        
        .Filters.Add "Access Databases", "*.md*;*.accd*"
        ' access databases
        
        .Filters.Add "Text Files", "*.csv;*.txt;*.log"
        ' text files
        
        .FilterIndex = typFileTypeIndex
        ' set filter index as requested by origin function
        
    End If
    
    If strTitle$ <> vbNullString Then
    ' if a title is provided by the origin function, ...
    
        .Title = strTitle$
        ' set the title
        
    End If
    
    .InitialFileName = strInitialFilePath$
    ' set initial file path, if provided by origin function

    If .Show = -1 Then
    ' if the user did not hit cancel, ...

        Dim arrSelected()
        ' variable for selected file(s)

        ReDim arrSelected(.SelectedItems.Count - 1)
        ' redim the selected file(s) array to the amount selected

        For lngCount& = 1 To .SelectedItems.Count
        ' for each selected item, ...
        
            arrSelected(lngCount& - 1) = .SelectedItems(lngCount&)
            ' update the selected file(s) array with the selected file(s)
            
        Next lngCount
        ' continue looping through each selected item
    
        arrResult(0) = True
        ' show success
        
        arrResult(1) = arrSelected()
        ' include selected file(s) in result
        
        FileDialog = arrResult()
        ' return result to origin function
        
    Else
    ' if the user did hit cancel (0), ...
    
        arrResult(0) = False
        ' show failure
        
        FileDialog = arrResult()
        ' return result to origin function

    End If

End With

End Function
