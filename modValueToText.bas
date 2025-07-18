Attribute VB_Name = "modValueToText"
' updated 2011.07.20
' created by Chris Staines

Public Sub ValueToText(shtTarget As Worksheet, lngColumn As Long, Optional boolHeader As Boolean)
' Converts a column's cell values to that column's
' displayed text.
'
' required:
'           shtTarget, the worksheet to be updated
'           lngColumn, the column of shtTarget to be updated
' optional:
'           boolHeader, whether to account for a header (start on line 2)

Dim lngRow As Long ' variable for current row in loop
Dim lngStart As Long ' variable for use with boolHeader

' disable Excel screen updating to speed up macro
Application.ScreenUpdating = False

If boolHeader = True Then _
    lngStart& = 2 Else lngStart& = 1

With shtTarget

' start on row 2 to avoid headers
    For lngRow& = lngStart& To .UsedRange.Rows.Count
    
        ' make the cell's value equal the cell's displayed text
        .Cells(lngRow&, lngColumn&).Value = .Cells(lngRow&, lngColumn&).Text
    
    Next lngRow&
    
End With

' re-enable screen updating, as macro is finished
Application.ScreenUpdating = True

End Sub
