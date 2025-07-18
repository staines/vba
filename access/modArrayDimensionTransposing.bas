Attribute VB_Name = "modArrayDimensionTransposing"
' updated 2016-10-31
' created by chris staines

Public Function ArrayDimensionTransposing(arrArray As Variant, intIndexToTranspose As Integer, Optional boolSkip0 As Boolean) As Variant
' takes a multidimension array and transpose 1 index into a single-dimension array

    Dim arrTemp()
    ' variable for temporary array, accounting for skipping
    
    If IsArray(arrArray) = True Then
    ' if an array, ...

        Dim intStart As Integer
        ' variable for where to start in the array
        
        If boolSkip0 = True Then intStart% = 1 Else intStart% = 0
        ' if skipping 1st, then skip

        ReDim arrTemp(0 To UBound(arrArray) - intStart%)
        ' redim the temporary array to support the necessary size
        
        Dim lngIndex As Long
        ' variable for iterating through the array
    
        For lngIndex& = intStart% To UBound(arrArray)
        ' from the start of, to the end of, the array, ...
        
            arrTemp(lngIndex& - intStart%) = arrArray(lngIndex&, intIndexToTranspose%)
            ' set temporary array index, accounting for skipping, as index of provided array
        
        Next lngIndex&
        ' continue to the next index
    
        ArrayDimensionTransposing = arrTemp
        ' return array to origin function
        
    ElseIf VarType(arrArray) = vbString Then
    ' if not an array, but a string, ...
    
        ReDim arrTemp(0 To 0)
        ' redim the temporary array to support a single index value
    
        arrTemp(0) = arrArray
        ' set temporary array index 0 as value
    
        ArrayDimensionTransposing = arrTemp
        ' return provided detail to origin function
    
    End If

End Function
