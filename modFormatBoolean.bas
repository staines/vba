Attribute VB_Name = "modFormatBoolean"
' convert boolean to desired string format
' 2016-07-18 chris staines

Option Compare Database

Public Enum BooleanFormat

    typYN = 1
    
    typYesNo = 2
    
    typTF = 3
    
    typTrueFalse = 4
    
    typNegativeOneZero = 5

End Enum

Public Function FormatBoolean(boolBoolean As Boolean, typFormat As BooleanFormat) As String
' convert a boolean to desired string format

    Select Case typFormat
    ' based on the desired format, ...
    
        Case 1
        ' y/n
        
            If boolBoolean = True Then FormatBoolean = "Y" Else FormatBoolean = "N"
            ' return desired format to origin function
    
        Case 2
        ' yes/no
        
            If boolBoolean = True Then FormatBoolean = "Yes" Else FormatBoolean = "No"
            ' return desired format to origin function
    
        Case 3
        ' t/f
        
            If boolBoolean = True Then FormatBoolean = "T" Else FormatBoolean = "F"
            ' return desired format to origin function
    
        Case 4
        ' true/false
        
            If boolBoolean = True Then FormatBoolean = "True" Else FormatBoolean = "False"
            ' return desired format to origin function
    
        Case 5
        ' -1/0
        
            If boolBoolean = True Then FormatBoolean = "0" Else FormatBoolean = "-1"
            ' return desired format to origin function
            
    End Select
    
End Function
