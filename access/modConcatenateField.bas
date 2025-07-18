Attribute VB_Name = "modConcatenateField"
Option Compare Database
' updated 2015.02.10
' created by Chris Staines
'
' dependencies:
'   modIsTable
'   modIsTableField

Public Enum FieldType
' field type enumerator

    ftText = 0
    ' field type of text
    
    ftNumber = 1
    ' field type of number

    ftDate = 2
    ' field type of date
    
End Enum

Public Function QuoteString(strString As String) As String
' return single quote or double quote for a given string, depending on existence of apostrophe in given string

If InStr(strString$, "'") > 0 Then QuoteString = Chr(34) & strString$ & Chr(34) Else QuoteString = "'" & strString$ & "'"
' return quoted string

End Function

Public Function ConcatenateField(strSource_Recordset As String, strSource_Field As String, ftType As FieldType, Optional boolTextDoubleQuote As Boolean) As String
' concatenate the contents of each row of a given recordset based on desired format
' not including initial check for if source recordset has data, as expected functionality may be to have no results based on empty source recordset

If IsTableField(strSource_Recordset$, strSource_Field$) = True Then
' if the given source table and field exists, ...

    Dim rsSource As Recordset
    ' variable for source table recordset

    Set rsSource = CurrentDb.OpenRecordset(strSource_Recordset$, dbOpenDynaset)
    ' open the source table recordset
    
    Dim strIn As String
    ' variable for in operator
    
    If Not (rsSource.BOF And rsSource.EOF) Then
    ' if not at the beginning and end of the source recordset (it's not empty), ...
    
        Do Until rsSource.EOF
        ' until the end of the source table, ...
        
            Select Case ftType
            ' based on the expected field type, ...
            
                Case ftText
                ' if text, ...
                
                    If boolTextDoubleQuote = False Then
                    ' if should use single quote (apostrophe) for text, ...
                    
                        If strIn$ = vbNullString Then _
                            strIn$ = QuoteString(rsSource.Fields(strSource_Field$).Value) _
                            Else strIn$ = strIn$ & ", " & QuoteString(rsSource.Fields(strSource_Field$).Value)
                        ' set the in operator string based on if adding to or creating new in operator string
                        
                    Else
                    ' if should use double quote for text, ...
                    
                        If strIn$ = vbNullString Then _
                            strIn$ = Chr(34) & rsSource.Fields(strSource_Field$).Value & Chr(34) _
                            Else strIn$ = strIn$ & ", " & Chr(34) & rsSource.Fields(strSource_Field$).Value & Chr(34)
                        ' set the in operator string based on if adding to or creating new in operator string
                        
                    End If

                Case ftNumber
                ' if number, ...
                
                    If strIn$ = vbNullString Then _
                        strIn$ = rsSource.Fields(strSource_Field$).Value _
                        Else strIn$ = strIn$ & ", " & rsSource.Fields(strSource_Field$).Value
                    ' set the in operator string based on if adding to or creating new in operator string
                
                Case ftDate
                ' if date, ...
                
                    If strIn$ = vbNullString Then _
                        strIn$ = "#" & rsSource.Fields(strSource_Field$).Value & "#" _
                        Else strIn$ = strIn$ & ", #" & rsSource.Fields(strSource_Field$).Value & "#"
                    ' set the in operator string based on if adding to or creating new in operator string
            
            End Select
        
            rsSource.MoveNext
            ' move to the next record in the recordset
        
        Loop
        ' continue to the next record in the recordset
        
    End If

    rsSource.Close
    ' close the source recordset

    Set rsSource = Nothing
    ' clear from memory

    ConcatenateField = strIn$
    ' return success
    
Else
' if the given source table does not exist, ...

    ConcatenateField = vbNullString
    ' return blank

End If

End Function
