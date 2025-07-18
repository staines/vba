Attribute VB_Name = "modIsTableField"
Option Compare Database
' updated 2014.10.21
' created by Chris Staines
'
' determine if a field exists for a provided table
'
' dependencies:
'   modIsTable

Public Function IsTableField(strTable As String, strField As String) As Boolean
' return true/false for if a field exists in a provided table

If IsTable(strTable$) = True Then
' if the provided table name is a table, ...

    Dim rsTable As Recordset
    ' variable for recordset
    
    Set rsTable = CurrentDb.OpenRecordset(strTable$, dbOpenTable)
    ' open the provided table as a recordset
    
    Dim objField As Field
    ' variable for field object
    
    For Each objField In rsTable.Fields
    ' for each field in the table, ...
    
        If objField.Name = strField Then IsTableField = True: Exit For
        ' if the field name is the provided field, return success
        
    Next objField
    ' continue to the next field
    
    rsTable.Close
    ' close the table
    
    Set rsTable = Nothing
    ' clear from memory

End If

End Function
