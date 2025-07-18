Attribute VB_Name = "modDropTableMatching"
Option Compare Database
' updated 2014.05.15
' created by Chris Staines
'
' note:  pick either of the two functions, which do the same thing (through different means).

Public Function DropTableMatching_TableDef(strTableMatchString As String) As Boolean
' drop any table with a name that matches the provided (case-insensitive, LIKE-compatible) string
'
' criteria examples:
'   "*hello*" would match on tables containing, "hello," anywhere in the name
'   "*hello" would match on tables with, "hello," at the end of the name
'   "hello*" would match on tables with, "hello," at the beginning of the name
'   "hel*lo" would match on tables with, "hel," at the beginning and, "lo," at the end of the name
'   "*hel*lo*" would match on tables with, "hel," and, "lo," anywhere in the name (though not where, "lo," preceeds, "hel")

Dim tdTable As TableDef
' variable for table to work with

For Each tdTable In CurrentDb.TableDefs
' for each table in the current database, ...

    'If tdTable.Name = "S1PR1_S_BU" Then Debug.Print tdTable.Attributes
    Select Case tdTable.Attributes
    ' based on the table attributes, ...
    
        Case 0, 1048576, 536870912
        ' if a local table or external table (but not a system object), ...

            If tdTable.Name Like strTableMatchString$ Then DoCmd.SetWarnings False: DoCmd.RunSQL "DROP TABLE [" & tdTable.Name & "];": DoCmd.SetWarnings True
            ' if the table name matches the provided string, drop the table
        
    End Select

Next tdTable
' continue to the next table

Set tdTable = Nothing
' clear from memory

DropTablesMatching_TableDefs = True
' return success to origin function (success due to no errors-- not that a table was found and dropped)

End Function

Public Function DropTableMatching_SQL(strTableMatchString As String) As Boolean
' drop any table with a name that matches the provided (case-insensitive, LIKE-compatible) string
'
' criteria examples:
'   "*hello*" would match on tables containing, "hello," anywhere in the name
'   "*hello" would match on tables with, "hello," at the end of the name
'   "hello*" would match on tables with, "hello," at the beginning of the name
'   "hel*lo" would match on tables with, "hel," at the beginning and, "lo," at the end of the name
'   "*hel*lo*" would match on tables with, "hel," and, "lo," anywhere in the name (though not where, "lo," preceeds, "hel")

Dim rsMatch As Recordset
' variable for recordset

Set rsMatch = CurrentDb.OpenRecordset("SELECT MSysObjects.Name FROM MSysObjects WHERE (MSysObjects.Type In (1, 4)) AND (MSysObjects.Flags In (0, 1048576)) AND (MSysObjects.Name LIKE '" & strTableMatchString$ & "');", dbOpenDynaset)
' open the recordset

With rsMatch
' working with the recordset, ...

    If Not (.EOF And .BOF) Then
    ' if not at the end and the beginning of the recordset (the recordset isn't empty), ...
    
        .MoveFirst
        ' move to the first record in the recordset
        
        Do Until .EOF
        ' until the end of the recordset, ...
        
            DoCmd.SetWarnings False: DoCmd.RunSQL "DROP TABLE [" & .Fields("Name") & "];": DoCmd.SetWarnings True
            ' if the table name matches the provided string, drop the table
            
            .MoveNext
            ' move to the next record in the recordset
        
        Loop
        ' attempt to loop to the next record in the recordset
    
    End If
    
    .Close
    ' close the recordset
    
End With
'Debug.Print rsMatch.RecordCount

Set rsMatch = Nothing
' clear from memory

End Function
