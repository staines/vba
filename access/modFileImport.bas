Attribute VB_Name = "modFileImport"
Option Compare Database
' updated 2015.08.31
' created by Chris Staines

Function FileImport(strFile As String, boolClearTable As Boolean, strTableName As String, Optional varFormatRowArray, Optional boolMessageOnError As Boolean, Optional strOriginalFilename As String, Optional boolSkipFormatInExcel As Boolean) As Boolean
' open, format, save, and import a file

If Dir(strFile$) <> Right(strFile$, Len(strFile$) - InStrRev(strFile$, "\")) Then FileImport = False: Exit Function
' if file does not exist, exeunt

If boolMessageOnError = True Then
' if user wishes to receive prompts on errors, ...

    Dim strFilenameForFeedback As String
    ' variable for filename in feedback
    
    If strOriginalFilename$ <> vbNullString Then
    ' if an original filename is provided, ...
    
        strFilenameForFeedback$ = strOriginalFilename$
        ' set the filename for feedback as the original filename
        
    Else
    ' if an original filename is not provided, ...
    
        strFilenameForFeedback$ = Right(strFile$, Len(strFile$) - InStrRev(strFile$, "\"))
        ' set the filename for feedback as the file being imported
    
    End If
    
End If

If boolClearTable = True Then
' if user wishes to clear the existing entries in the table, ...

    DoCmd.SetWarnings False
    ' disable user prompts

    DoCmd.RunSQL "DELETE * FROM [" & strTableName$ & "];"
    ' clear the table
    
    DoCmd.SetWarnings True
    ' re-enable user prompts
    
End If

If boolSkipFormatInExcel = False Then
' if user wishes to format the workbook, ...
    
    Dim xlApp As Excel.Application
    ' variable for excel application
    
    Dim xlBook As Excel.Workbook
    ' variable for excel workbook
    
    Set xlApp = CreateObject("Excel.Application")
    ' work with excel to clean the file
    
    xlApp.Visible = True
    ' show the excel window
    
    Set xlBook = xlApp.Workbooks.Open(strFile$, True, False, , , , , , , , False)
    ' open the file to clean
    
    'xlApp.ScreenUpdating = False
    ' disable screen updating to speed up processing of file
    
    If Rank(varFormatRowArray) >= 2 Then
    ' if a format row array was provided and is the right size (at least 2 dimensions), ...
    
        Dim lngRowIndex As Long
        ' variable for index of row being formatted
        
        lngRowIndex& = 1
        ' default row index to 1
        
        Dim lngRowDeletedCount As Long
        ' variable for number of rows deleted
        
        Dim lngFormatRowArray_Index As Long
        ' variable for index in format row array
        
        For lngFormatRowArray_Index& = 0 To UBound(varFormatRowArray, 2)
        ' from the 1st index of the format row array, to the last index, ...
    
            If varFormatRowArray(lngFormatRowArray_Index&, 1) = True Then
            ' if the rows should be deleted, ...
    
                Do
                ' attempt to modify the requested amount of rows for the specified amount...
    
                    xlBook.Sheets(1).Rows(lngRowIndex& & ":" & lngRowIndex&).Select
                    ' select the header row
                    
                    xlApp.Selection.Delete Shift:=xlUp
                    ' delete the extraneous header row
                    
                    lngRowDeletedCount& = lngRowDeletedCount + 1
                    ' iterate to the next row index
                    
                Loop Until lngRowDeletedCount& = varFormatRowArray(lngFormatRowArray_Index&, 0)
                ' continue looping through until the requested number of rows
                
                lngRowDeletedCount& = 0
                ' reset row deleted count
    
            Else
            ' if the rows should not be deleted, ...
            
                lngRowIndex& = lngRowIndex& + varFormatRowArray(lngFormatRowArray_Index&, 0)
                ' add the requested number of rows to the row index
            
            End If
                
        Next lngFormatRowArray_Index&
        ' continue to the next index
        
    End If
    
    Dim rngFind As Range
    ' variable for range to find what
    
    xlBook.Sheets(1).Rows("1:1").Select
    ' select the first row
    
    Set rngFind = xlApp.Selection.Find(What:=".", LookAt:=xlPart, SearchOrder:=xlByColumns, SearchDirection:= _
        xlNext, MatchCase:=False, SearchFormat:=False)
    ' attempt to find the what in the selection
    
    If Not rngFind Is Nothing Then
    ' if the what was found (it's not nothing), ...
    
        xlApp.Selection.Replace What:=".", Replacement:="", LookAt:=xlPart, _
            SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
            ReplaceFormat:=False
        ' replace what with replacement in header row
        
    End If
    
    Set rngFind = Nothing
    ' clear from memory
    
    xlApp.ScreenUpdating = True
    ' re-enable screen updating
    
    Dim strFileName_Temporary As String
    ' variable for temporary filename
    
    strFileName_Temporary$ = Left(strFile$, Len(strFile$) - (Len(strFile$) - InStrRev(strFile$, ".")) - 1) & "_temp.xlsx"
    ' derive temporary file name from current file name
    
    xlBook.SaveAs FileName:= _
        strFileName_Temporary$, _
        FileFormat:=xlOpenXMLWorkbook, _
        CreateBackup:=False, _
        ConflictResolution:=xlLocalSessionChanges
    ' save the file in a more agreeable format
    
    xlBook.Close True
    ' save and close the workbook
    
    Set xlBook = Nothing
    ' clean up for memory
    
    xlApp.Quit
    ' exit excel
    
    Set xlApp = Nothing
    ' clean up for memory
    
    DoCmd.SetWarnings False
    ' disable user prompts
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, strTableName$, strFileName_Temporary$, True
    ' import the selected file
    
    DropTableMatching_TableDef "*ImportErrors*"
    ' drop any import errors table
    
    Kill strFileName_Temporary$
    ' kill the temporary file
    
    FileImport = True
    ' return success to origin function

Else
' if user does not wish to format the workbook, ...
    
    DoCmd.SetWarnings False
    ' disable user prompts
    
    DoCmd.TransferSpreadsheet acImport, acSpreadsheetTypeExcel12Xml, strTableName$, strFile$, True
    ' import the selected file
    
    DropTableMatching_TableDef "*ImportErrors*"
    ' drop any import errors table

    FileImport = True
    ' return success to origin function

End If

End Function
