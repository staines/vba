Attribute VB_Name = "modCopyWorksheetsToWorkbook"
' updated 2016-10-31
'
' created by chris staines
'
' copy worksheets to a new workbook, with extra options for formulas

Public Function CopyWorksheetsToWorkbook(strWorkbookFile As String, arrSheets As Sheets, _
    Optional boolReplaceFormulae As Boolean, Optional arrReplaceFormulaeSheets As Variant, Optional arrReplaceFormulae As Variant) As Boolean
' copy the provided array of sheets to a new workbook, optionally replacing formula-containing cells with cell values

    arrSheets.Copy
    ' copy sheets to the new workbook
    
    Dim wbkCopiedTo As Workbook
    ' variable for workbook to work with
    
    Set wbkCopiedTo = ActiveWorkbook
    ' set new/active workbook as workbook to work with
    
    If boolReplaceFormulae = True Then
    ' if replacing formulae, ...
    
        Dim shtTemp As Worksheet
        ' variable for temporary worksheet(s) to replace formulae in
        
        Dim lngFormulaeIndex As Long
        ' variable for iterating through formulae variable
        
        Dim rngFind As Range
        ' variable for find range
        
        Dim rngFormulaCell As Range
        ' variable for formula cell range
    
        For Each shtTemp In wbkCopiedTo.Sheets(arrReplaceFormulaeSheets)
        ' for each worksheet in which the formulae should be replaced, ...

            For lngFormulaeIndex& = 0 To UBound(arrReplaceFormulae)
            ' from the first formula, to the last formula, ...

                Set rngFind = shtTemp.Cells.Find(What:="=" & arrReplaceFormulae(lngFormulaeIndex&), LookIn:=xlFormulas)
                ' looking for the formula, ...
                
                If Not rngFind Is Nothing Then
                ' if the formula is found, ...

                    Do
                    ' for every time the formula is found, ...
                    
                        rngFind.Value = rngFind.Value
                        ' set the cell value as the value, without formula
                        
                        'rngFind.Copy
                        ' copy the cell value
                        
                        'rngFind.PasteSpecial xlPasteValuesAndNumberFormats
                        ' paste the cell value
                        
                        Set rngFind = shtTemp.Cells.FindNext(rngFind)
                        ' attempt to find the next instance of the formula
                    
                    Loop While Not rngFind Is Nothing
                    ' loop to the next range, if found

                End If
                    
            Next lngFormulaeIndex&
            ' continue to the next formula
        
        Next shtTemp
        ' continue to next worksheet
        
    End If

    Application.DisplayAlerts = False
    ' disable user prompts
    
    wbkCopiedTo.SaveAs Filename:=strWorkbookFile$, FileFormat:=xlExcel7, ConflictResolution:=xlLocalSessionChanges
    ' save workbook
    
    wbkCopiedTo.Close 'SaveChanges:=True, Filename:=strWorkbookFile$
    ' save the workbook
    
    Application.DisplayAlerts = True
    ' re-enable user prompts
    
    CopyWorksheetsToWorkbook = True
    ' return success to origin function
    ' only a catastrophic failure would stop this code, so we'll always return true :/
    
End Function


