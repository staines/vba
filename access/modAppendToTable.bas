Attribute VB_Name = "modAppendToTable"
Option Compare Database

' updated 2014.10.21
' created by Chris Staines
'
' attempt to append data to a table; return relevant errors if unable to
'
' dependencies:
'   modIsTable
'   modRank

Public Function AppendToTable(strTable As String, varField As Variant) As Variant
' attempt to append data to a table; return relevant errors if unable to
'
' format:
'   strtable = table name to append value to
'   varfield = array (1,index) containing field name and value to append to table
'       0 = field name
'       1 = value
'
' result:
'   0 = success true/false
'   1 = reason if unsuccessful

Dim varResult(1)
' variable for result

If IsTable(strTable$) = True Then
' if the table name provided is a table, ...

    If Rank(varField) > 1 Then
    ' if the rank of the provided field array is 2 or more, ...

        Dim rsTable As Recordset
        ' variable for recordset
        
        Set rsTable = CurrentDb.OpenRecordset(strTable$, dbOpenTable)
        ' open the provided table as a recordset
        
        rsTable.AddNew
        ' set intention to add a new record
  
        Dim lngFieldIndex As Long
        ' variable for index in field array
        
        lngFieldIndex& = 0
        ' set to 0
        
        Dim objField As Field
        ' variable for field object
        
        Dim boolFieldFound As Boolean
        ' variable for if the field was found in the table

        Do While lngFieldIndex& <= UBound(varField, 2)
        ' while the index in the field array is less than or equal to the highest index in the 2nd dimension of the array, ...
        
            boolFieldFound = False
            ' reset the field found boolean, so we can see if the current field is found
        
            For Each objField In rsTable.Fields
            ' for each field in the table, ...

                If objField.Name = varField(0, lngFieldIndex&) Then objField.value = varField(1, lngFieldIndex&): boolFieldFound = True: Exit For
                ' if the name of the field matches the field array field name, set the field value accordingly
                
            Next objField
            ' continue to the next field in the table
            
            If boolFieldFound = False Then varResult(0) = False: varResult(1) = "Provided field not found in table.": rsTable.CancelUpdate: Exit Do
            ' if the field was not found, set result, cancel the new record, and exit the do

            lngFieldIndex& = lngFieldIndex& + 1
            ' increment field array index

        Loop
        ' continue the loop

        If varResult(0) = vbNullString Then
        ' if the result has not been set as false, ...

            rsTable.Update
            ' update the table
        
            varResult(0) = True
            ' return success

        End If
        
        rsTable.Close
        ' close the table
        
        Set rsTable = Nothing
        ' clear from memory
        
    Else
    ' if the rank of the provided field array is not the correct size or larger, ...
    
        varResult(0) = False
        ' return failure
        
        varResult(1) = "Provided field array is not the correct size."
        ' provide reason for failure
    
    End If
    
Else
' if the table provided is not a table, ...

    varResult(0) = False
    ' return failure
    
    varResult(1) = "Provided table is not a recognized table."
    ' provide reason for failure

End If

AppendToTable = varResult
' return result to origin function

End Function

