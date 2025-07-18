Attribute VB_Name = "modFormatWorkbook"
' updated 2015.01.15
'
' created by Chris Staines
'
' normalize the formatting of all worksheets in a workbook

Public Function FormatWorkbook(strFile As String, Optional boolFormatAsTable As Boolean, Optional boolColumnAutoFit As Boolean) As Boolean
' format a given workbook based on filename

If Dir(strFile$) <> vbNullString Then
' if the file exists, ...
    
    Dim xlApp As Excel.Application
    ' variable for excel application
    
    Dim xlBook As Excel.Workbook
    ' variable for excel workbook
    
    Set xlApp = CreateObject("Excel.Application")
    ' open an instance of excel
    
    xlApp.Visible = True
    ' show the excel window
    
    Set xlBook = xlApp.Workbooks.Open(strFile$, True, False, , , , , , , , False)
    ' open the given file
    
    Dim xlSheet As Excel.Worksheet
    ' variable for interacting with a worksheet instance
    
    For Each xlSheet In xlBook.Worksheets
    ' for each sheet in the workbook, ...
        
        xlSheet.Activate
        ' activate the sheet to avoid excel propensity to reference active sheet
        
        xlSheet.ListObjects.Add(xlSrcRange, _
            xlSheet.Range("$A$1:" & xlSheet.Cells(xlSheet.UsedRange.Rows.Count, xlSheet.UsedRange.Columns.Count).Address & ""), _
            , _
            xlYes).Name = _
                "Table_" & xlSheet.Name
        ' format the entire sheet as a table
        
        xlSheet.Cells.Select
        ' select all cells in the sheet
        
        xlSheet.Cells.EntireColumn.AutoFit
        ' autofit each column in the sheet
        
        xlSheet.Range("A1").Select
        ' change selection to first cell
                
        'Debug.Print xlSheet.Cells(xlSheet.UsedRange.Rows.Count, xlSheet.UsedRange.Columns.Count).Address
        
    Next xlSheet
    ' continue to the next sheet in the workbook
    
    xlBook.Close True
    ' close workbook and save if needed
    
    Set xlBook = Nothing
    ' clear from memory
    
    xlApp.Quit
    ' exit excel
    
    Set xlApp = Nothing
    ' clear from memory
    
    FormatWorkbook = True
    ' return success to origin function
    
Else
' if the file does not exist, ...

    FormatWorkbook = False
    ' return failure to origin function
    
End If

End Function
