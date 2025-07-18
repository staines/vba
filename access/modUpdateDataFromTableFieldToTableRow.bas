Attribute VB_Name = "modUpdateDataFromTableFieldToTableRow"
Option Compare Database
'
' created by Chris Staines

Public Sub UseTableFieldNameToReplaceStringAndRunQuery(strTableWithField As String, strQueryToRun As String, _
    strReplaceString As String, Optional strReplaceStringAndEscapeDoubleQuotes As String, Optional arrFieldsToIgnore)
' use the field name(s) of a table to replace a string in a query and run the given query

    Dim rsTableWithField As Recordset
    ' variable for recordset

    Set rsTableWithField = CurrentDb.OpenRecordset(strTableWithField, dbOpenTable)
    ' open the recordset

    If Not (rsTableWithField.BOF And rsTableWithField.BOF) Then
    ' if not at the beginning and end of the field table (it's not empty), ...
    
        Dim objField As Field
        ' variable for field object
        
        For Each objField In rsTableWithField.Fields
        ' for each field in the table, ...
        
            If IsArray(arrFieldsToIgnore) Then
            ' if fields to ignore have been provided, ...
            
                If Not UBound(Filter(arrFieldsToIgnore, objField.Name)) > -1 Then
                ' if the field name is not in the to ignore array, ...

                    DoCmd.SetWarnings False
                    ' disable user prompts

                    DoCmd.RunSQL Replace(Replace(strQueryToRun$, strReplaceStringAndEscapeDoubleQuotes$, Replace(objField.Name, """", """""")), strReplaceString$, objField.Name)
                    ' run the provided sql and replace string(s) as necessary

                    DoCmd.SetWarnings True
                    ' re-enable user prompts
                
                End If
            
            Else
            ' if fields to ignore have not been given, ...

                DoCmd.SetWarnings False
                ' disable user prompts

                DoCmd.RunSQL Replace(Replace(strQueryToRun$, strReplaceStringAndEscapeDoubleQuotes$, Replace(objField.Name, """", """""")), strReplaceString$, objField.Name)
                ' run the provided sql and replace string(s) as necessary

                DoCmd.SetWarnings True
                ' re-enable user prompts
                
            End If

        Next objField
        ' continue to the next field
    
    End If
    
    rsTableWithField.Close
    ' close the recordset
    
    Set rsTableWithField = Nothing
    ' clear the recordset from memory

End Sub
