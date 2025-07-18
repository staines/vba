Attribute VB_Name = "modIsProperty"
Option Compare Database
' updated 2014.10.09
' created by Chris Staines
'
' requires modIsTable

Function IsProperty(strTable As String, strProperty As String) As Boolean
' return true/false for if a table property exists

If IsTable(strTable$) = True Then
' if the provided table exists, ...
    
    Dim rsTable As Recordset
    ' variable for table recordset

    Set rsTable = CurrentDb.OpenRecordset(strTable$, dbOpenDynaset)
    ' open the table
    
    Dim objProperty As Property
    ' variable for property object
    
    For Each objProperty In rsTable.Properties
    ' for each table property, ...
    
        If objProperty.Name = strProperty$ Then IsProperty = True: Exit For
        ' if the table property name matches the provided property, return success
        
    Next objProperty
    ' continue to the next table property
    
    rsTable.Close
    ' close the table
    
    Set objProperty = Nothing
    ' clear from memory
    
    Set rsTable = Nothing
    ' clear from memory
    
End If

End Function

Function SetProperty(strTable As String, strPropertyName As String, varPropertyValue As Variant, dbPropertyType As DataTypeEnum, _
    Optional boolCreatePropertyIfDoesNotExist As Boolean) As Boolean
' set a table property, creating it if necessary

If IsTable(strTable$) = True Then
' if the table provided exists, ...

    Dim rsTable As Recordset
    ' variable for table recordset

    Set rsTable = CurrentDb.OpenRecordset(strTable$, dbOpenDynaset)
    ' open the table
    
    Dim objProperty As Property
    ' variable for property object
    
    For Each objProperty In rsTable.Properties
    ' for each table property, ...
    
        If objProperty.Name = strPropertyName$ Then
        ' if the table property name matches the provided property, ...
        
            objProperty.value = varPropertyValue
            ' set the property value
        
            SetProperty = True
            ' return success!
            
            Exit For
            ' exit the for loop
            
        End If
        
    Next objProperty
    ' continue to the next table property
    
    rsTable.Close
    ' close the table
    
    Set objProperty = Nothing
    ' clear from memory
    
    Set rsTable = Nothing
    ' clear from memory
    
    If SetProperty = False And boolCreatePropertyIfDoesNotExist = True Then
    ' if the property did not exist and it should be created, ...

        Dim dBase As DAO.Database
        ' variable for dao database
        
        Dim tblDef As DAO.TableDef
        ' variable for tabledef

        Set dBase = CurrentDb
        ' set the dao database as the current database
        
        Set tblDef = dBase.TableDefs(strTable$)
        ' set the tabledef to work with as the provided table
         
        Set objProperty = tblDef.CreateProperty(strPropertyName$, dbPropertyType, varPropertyValue)
        ' create the new property in the tabledef
        
        tblDef.Properties.Append objProperty
        ' append the new property to the tabledef properties
        
        tblDef.Properties.Refresh
        ' refresh the tabledef properties
        
        If IsProperty(strTable$, strPropertyName$) = True Then SetProperty = True
        ' if the property now exists, return success to the origin function
    
    End If
    
End If

End Function
