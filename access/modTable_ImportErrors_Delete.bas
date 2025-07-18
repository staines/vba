Attribute VB_Name = "modTable_ImportErrors_Delete"
' updated 2016.07.18
' created by Chris Staines
Option Compare Database

Public Function Table_ImportErrors_Delete()
' delete any *_importerrors tables
    
    Dim objTable As TableDef
    ' variable for looping through tables
    
    For Each objTable In CurrentDb.TableDefs
    ' loop through all tables, ...
    
        If objTable.Name Like "*_ImportErrors*" Then
        ' find those which are import errors, ...
        
            DoCmd.SelectObject acTable, objTable.Name, True
            ' select them, ...
            
            DoCmd.DeleteObject acTable, objTable.Name
            ' and delete them.
            
        End If
        
    Next objTable
    ' continue looping through all tables

End Function
