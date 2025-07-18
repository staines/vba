Attribute VB_Name = "modNormalize"
' updated 2014.01.06
' created by Chris Staines

Option Compare Database

Public Enum NormalizationType
' enumerator for normalize

    typDate = 0
    typCurrency = 1
    typString = 2
    typHTML = 3
    
End Enum

Public Enum NormalizationStringType
' enumerator for normalize

    typAlphanumeric = 0
    typAlphanumericWithSpace = 1
    typAlphaWithSpace = 2
    typAlphaOnly = 3
    typNumericOnly = 4
    
End Enum

Public Function Normalize(strString As String, intType As NormalizationType, Optional boolHTMLSpaceAsPlus As Boolean, Optional typNormalizationStringType As NormalizationStringType)
' normalize string by desired type
'
' types:
'       typDate:        normalize date string to date (handles yyyymmdd gracefully)
'       typCurrency:    normalize currency string (handles null values)
'       typString:      normalize string to remove all characters except A to Z, a to z, and 0 to 9
'       typHTML:        normalize string, encoding each non-alphanumerical character to the html equivalent

Select Case intType
' handle each type differently

    Case 0, typDate
    ' dates
        
        strString$ = Trim(strString$)
        ' trim string to remove misleading spaces

        Dim datTemp As Date
        ' temporary variable for storing dates

        If strString$ = vbNullString Or strString$ = Null Then
        ' if there is no date, act accordingly
        
            Select Case App
            ' based on which app we're working in, ...
            
                Case vbAccess
                ' if we're in Access, ...
                
                    Normalize = Null: Exit Function
                    ' Access prefers Null values for blank dates, whereas ...
                    
                Case Else
                ' if we're not in Access, ...
                
                    Normalize = vbNullString: Exit Function
                    ' excel prefers a blank string
                    
            End Select
            
        End If
        
        If Len(strString$) = 11 And InStr(strString$, "-") <> 0 Then
        ' if string is DD-MMM-YYYY, ...
        
            datTemp = Format(strString$, "mm/dd/yyyy")
            ' VBA handles DD-MMM-YYYY formatting gracefully
            
            Normalize = datTemp: Exit Function
            ' exit to avoid conflict
            
        End If
        
        For lngPosition = 1 To Len(strString$)
        ' go through the string, 1 character at a time

            If IsNumeric(Mid(strString$, lngPosition, 1)) = False And _
                Mid(strString$, lngPosition, 1) <> "/" Then
            ' if it's not a valid ##/##/#### date, then ...

                Select Case App
                ' based on which app we're working in, ...
            
                    Case vbAccess
                    ' if we're in Access, ...
                    
                        Normalize = Null: Exit Function
                        ' Access prefers Null values for blank dates, whereas ...
                        
                    Case Else
                    ' if we're not in Access, ...
                    
                        Normalize = vbNullString: Exit Function
                        ' we can use a blank string
                    
                End Select
                
            End If
            
        Next lngPosition
    
        If Len(strString$) = 8 And InStr(strString$, "/") = 0 Then
        ' if string is probably yyyymmdd, ...
        
            datTemp = Mid(strString$, 5, 2) & "/" & Right(strString$, 2) & "/" & Left(strString$, 4)
            ' set the date to mm/dd/yyyy
            
            Normalize = datTemp: Exit Function
            ' exit to avoid conflicts
            
        Else
        ' otherwise, if string is obviously not yyyymmdd, handle as a regular date
            
            datTemp = Format(strString$, "mm/dd/yyyy")
            ' otherwise, just format the date properly
            
            Normalize = datTemp: Exit Function
            ' exit to avoid conflicts
            
        End If
        
    Case 1, typCurrency
    ' currencies
        
        strString$ = Trim(strString$)
        ' trim string to remove misleading spaces

        If strString$ = vbNullString Then strString$ = "0"
        ' set null values to 0

        strString$ = FormatCurrency(strString$)
        ' easy use of formatcurrency
        
        Normalize = strString$: Exit Function
        ' we keep it a string to show $, etc. -- currency types are just doubles
        
    Case 2, typString
    ' strings
            
        strString$ = Trim(strString$)
        ' trim string to remove misleading spaces

        Dim intCharacter As Integer
        ' variable to store character location when looping through a string
        
        Dim strTemp As String
        ' temporary variable for the sanitized string
    
        strString$ = Trim(strString$)
        ' trim excess spaces, etc. from the string (just because we can)
        
        For intCharacter% = Len(strString$) To 1 Step -1
        ' from the end of the string to the (would be) first character, ...
    
            Select Case typNormalizationStringType
            ' working with the normalization string type, ...
            
                Case typAlphanumeric, 0
                ' if user wants alphanumeric only, ...
                
                    Select Case Asc(Mid(strString$, intCharacter%, 1))
                    ' working with the ascii index value of the individual character, ...
                    
                        Case 97 To 122, 65 To 90, 48 To 57
                        ' if the character is a to z, A to Z, or 0 to 9, ..
                            
                            strTemp$ = Mid(strString$, intCharacter%, 1) & strTemp$
                            ' add it to the beginning of the temporary string
                    
                    End Select
                    
                Case typAlphanumericWithSpace, 1
                ' if user wants alphanumeric and spaces only, ...
        
                    Select Case Asc(Mid(strString$, intCharacter%, 1))
                    ' working with the ascii index value of the individual character, ...
                    
                        Case 97 To 122, 65 To 90, 48 To 57, 32
                        ' if the character is a to z, A to Z, or 0 to 9, ..
                            
                            strTemp$ = Mid(strString$, intCharacter%, 1) & strTemp$
                            ' add it to the beginning of the temporary string
                    
                    End Select

                Case typAlphaWithSpace, 2
                ' if user wants alphas (letters) and spaces only, ...
        
                    Select Case Asc(Mid(strString$, intCharacter%, 1))
                    ' working with the ascii index value of the individual character, ...
                    
                        Case 97 To 122, 65 To 90, 32
                        ' if the character is a to z, A to Z, or 0 to 9, ..
                            
                            strTemp$ = Mid(strString$, intCharacter%, 1) & strTemp$
                            ' add it to the beginning of the temporary string
                    
                    End Select
                    
                Case typAlphaOnly, 3
                ' if user wants alphas (letters) only, ...
        
                    Select Case Asc(Mid(strString$, intCharacter%, 1))
                    ' working with the ascii index value of the individual character, ...
                    
                        Case 97 To 122, 65 To 90
                        ' if the character is a to z, A to Z, or 0 to 9, ..
                            
                            strTemp$ = Mid(strString$, intCharacter%, 1) & strTemp$
                            ' add it to the beginning of the temporary string
                    
                    End Select
                    
                
                Case typNumericOnly, 4
                ' if user wants alphanumeric and spaces only, ...
        
                    Select Case Asc(Mid(strString$, intCharacter%, 1))
                    ' working with the ascii index value of the individual character, ...
                    
                        Case 48 To 57
                        ' if the character is a to z, A to Z, or 0 to 9, ..
                            
                            strTemp$ = Mid(strString$, intCharacter%, 1) & strTemp$
                            ' add it to the beginning of the temporary string
                    
                    End Select
                    
            End Select
            
        Next intCharacter%
        ' continue looping through
        
        Normalize = strTemp$
        ' return result to origin function
        
    Case 3, typHTML
    ' html string
    ' partially derived from an example by tomalak on stackoverflow.com
    
        Dim strTempHTML As String
        ' variable for temporary html string
        
        Dim intCharacterHTML As Integer
        ' variable to store character location when looping through html string
        
        For intCharacter% = 1 To Len(strString$)
        ' from the first character to the last, ...
        
            Select Case Asc(Mid(strString$, intCharacter%, 1))
            ' working with the ascii index value of the individual character, ...

                Case 0 To 15
                ' if a control character, ...
                
                    strTemp$ = strTemp$ & "%0" & Hex(Asc(Mid(strString$, intCharacter%, 1)))
                    ' add the encoded version to the end of the temporary string
                    
                Case 32
                ' if a space, ...
                
                    If boolHTMLSpaceAsPlus = True Then
                    ' if the space should be a plus and not encoded, ...
                    
                        strTemp$ = strTemp$ & "+"
                        ' add a plus to the end of the temporary string
                        
                    Else
                    ' if the space should be encoded, ...
                    
                        strTemp$ = strTemp$ & "%" & Hex(Asc(Mid(strString$, intCharacter%, 1)))
                        ' add the encoded space to the end of the temporary string
                        
                    End If
                    
                Case 97 To 122, 65 To 90, 48 To 57
                ' if the character is a to z, A to Z, or 0 to 9, ..
                    
                    strTemp$ = strTemp$ & Mid(strString$, intCharacter%, 1)
                    ' add the character itself to the end of the temporary string
                    
                Case Else
                ' if an unrecognized character, ...
                
                    strTemp$ = strTemp$ & "%" & Hex(Asc(Mid(strString$, intCharacter%, 1)))
                    ' add the encoded character to the end of the temporary string
            
            End Select
            
        Next intCharacter%
        ' continue looping through
        
        Normalize = strTemp$
        ' return result to origin function
        
End Select

End Function
