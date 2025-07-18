Attribute VB_Name = "modListOpenSheets"
' created by Chris Staines

Public Sub ListOpenSheets(rngOutput As Excel.Range, Optional bookExclude As Excel.Workbook)
' provide an in-cell dropdown of open sheets for user selection
'
' required:
'           rngOutput, the range to put the dropdown in
' optional:
'           bookExclude, a book that may be excluded (such as the current one)

Dim bookTemp As Excel.Workbook ' for cycling through open workbooks
Dim strOpen As String ' string of open workbooks

For Each bookTemp In Excel.Workbooks

    If bookTemp.Name <> bookExclude.Name Then strOpen$ = strOpen$ & bookTemp.Name & ", "

Next bookTemp

If Not strOpen$ = vbNullString Then

    With rngOutput.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=strOpen$
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    rngOutput.Value = Left(strOpen$, InStr(strOpen$, ",") - 1)
    
Else

    MsgBox "Please open a[nother] workbook."
    
End If

End Sub
