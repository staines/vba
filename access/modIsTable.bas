Attribute VB_Name = "modIsTable"
' updated 2014.03.13
' created by Chris Staines
'
' determine if a table exists

Option Compare Database

Function IsTable(strTable As String) As Boolean
' return true/false for if a table exists

Dim tblIndex As TableDef
' variable of tabledef for loop

For Each tblIndex In CurrentDb.TableDefs
' for every tabledef in the database, ...

    If tblIndex.Name = strTable$ Then IsTable = True: Exit Function
    ' if the table name matches the provided name, show success and exeunt
    
Next tblIndex
' continue to the next tabledef

End Function
