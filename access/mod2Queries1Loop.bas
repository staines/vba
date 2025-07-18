Attribute VB_Name = "mod2Queries1Loop"
Option Compare Database
' updated 2013.07.16
' created by Chris Staines
'
' loop through a query/table, set a tempvar based on a field, and open a second query/table
'

Function TwoQueriesOneLoop(strTargetQuery As String, strSourceQuery As String, strSourceField As String, strTempVar As String)
' run a query based on the values in another query, utilizing a tempvar
' note:  skips null results from the source query

If strTargetQuery$ = vbNullString Then Exit Function
' exit if no target query provided

If strSourceQuery$ = vbNullString Then Exit Function
' exit if no source query provided

If strSourceField = vbNullString Then Exit Function
' exit if no source field provided

If strTempVar = vbNullString Then Exit Function
' exit if no tempvar provided

Dim rsSourceQuery As Recordset
' recordset to work with

Set rsSourceQuery = CurrentDb.OpenRecordset(strSourceQuery$, dbOpenDynaset)
' state intention to work with the source query

With rsSourceQuery
' working with the query, ...

    If Not (.EOF And .BOF) Then
    ' if the query is not empty, ...
    
        .MoveFirst
        ' move to the first record
    
        Do While Not .EOF
        ' while we are not at the end of the recordset, ...

            If Not IsNull(.Fields(strSourceField$).Value) Then
            ' if the value is not null, ...

                Access.TempVars.Item(strTempVar$).Value = .Fields(strSourceField$).Value
                ' set the temporary variable for the query as the dispatch provided
                
                DoCmd.OpenQuery strTargetQuery$
                ' open/run the target query

            End If
            
            .MoveNext
            
        Loop
        ' continue looping until the end of the recordset
        
    End If
        
End With

Set rsSourceQuery = Nothing
' clean up for memory

End Function


