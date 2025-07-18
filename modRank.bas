Attribute VB_Name = "modRank"
' created by Chris Staines
'
' obtain the rank of the array (number of dimensions)
' named after similar visual studio 2008 functionality

Public Function Rank(varArray As Variant) As Integer
' dimension rank in an array
' redone from http://support.microsoft.com/kb/152288

On Error Resume Next
' i ain't 'bout that on error life, but sometimes you need to hammer in a screw

If IsArray(varArray) = True Then
' if the array is provided, ...

    If IsEmpty(varArray) = False Then
    ' if the array is not empty, ...
    
        Dim lngDimension As Long
        ' variable for dimension in array
        
        For lngDimension& = 1 To 60000
        ' from the 1st to the maximum dimension, ...
        
            If IsNull(LBound(varArray, lngDimension&)) = True Then Rank = lngDimension& - 1: Exit For
            ' if the lbound for the dimension in the array is null,
            ' return the previous dimension to the origin function and exit the for
        
        Next lngDimension&
        ' continue to the next dimension (heh)
    
    End If
    
End If

End Function
