Attribute VB_Name = "modMailSend"
Option Compare Database
' updated 2014.03.26
' created by Chris Staines
'
' required reference for complete functionality:
'   Microsoft Outlook Object Library
'
' Portions from:
'   Ron de Bruin

Public Function Example()

Dim varFiles(1) As Variant
' variable for file array to send

varFiles(0) = Environ("USERPROFILE") & "\desktop\some local file.txt"
' first file to add

varFiles(1) = Environ("USERPROFILE") & "\desktop\another local file.txt"
' second file to add

Call mailSend("chris.c.staines@lowes.com", , , "Howdy, Subject!", "Happy I can send you an e-mail.", varFiles, True, False)
' send the e-mail

End Function

Public Sub mailSend(Optional strTo As String, Optional strCC As String, Optional strBBC As String, _
    Optional strSubject As String, Optional strBody As String, Optional varAttachmentFilepath As Variant, Optional boolSend As Boolean, _
    Optional boolHTMLBody As Boolean, Optional strSendFromEmail As String)
' send an e-mail message (and optional attachment) via outlook
' adapted from code by ron de bruin of www.rondebruin.nl

Dim objOutlook As Object
' object for outlook application

Dim objMessage As Object
' object for outlook mail message

Set objOutlook = CreateObject("Outlook.Application")
' reference/open outlook

Set objMessage = objOutlook.CreateItem(0)
' create a new outlook mail message

With objMessage
' working with the new outlook mail message, ...

    .To = strTo$
    ' set the recipients
    
    .CC = strCC$
    ' set the carbon copy recipients
    
    .BCC = strBCC$
    ' set the blind carbon copy recipients
    
    .Subject = strSubject$
    ' set the subject
    
    If boolHTMLBody = True Then
    ' if user requests to have the e-mail sent as an html e-mail, ...
    
        .BodyFormat = olFormatHTML
        ' set the body format of the e-mail to html
    
        .HTMLBody = strBody$
        ' set htmlbody of e-mail
        
    Else
    ' if user does not want to have the e-mail sent as an html e-mail, ...
    
        .Body = strBody$
        ' set the body
        
    End If
    
    If strSendFromEmail$ <> vbNullString Then .SentOnBehalfOfName = strSendFromEmail$
    ' if user provided a from e-mail, set the sentonbehalfof e-mail
    
    If IsArray(varAttachmentFilepath) Then
    ' if the attachment filepath is an array, ...
    
        If IsEmpty(varAttachmentFilepath) = False Then
        ' if the attachment filepath array is not empty, ...
    
            For Each objAttachmentFilepath In varAttachmentFilepath
            ' looping through each attachment filepath in the array, ...
            
                If CStr(objAttachmentFilepath) <> vbNullString Then
                ' if the attachment filepath, converted to a string, is not blank, ...
                
                    .Attachments.Add objAttachmentFilepath
                    ' add the attachment filepath to the array
                
                End If
                
            Next objAttachmentFilepath
            ' continue looping through the attachment filepath array
        
        End If
    
    Else
    ' if the attachment filepath is not an array, ...
    
        If IsMissing(varAttachmentFilepath) = False Then
        ' if an attachment filepath is possibly provided, ...
        
            If varAttachmentFilepath <> vbNullString Then
            ' if an attachment filepath is provided
        
                .Attachments.Add varAttachmentFilepath
                ' add the attachment
            
            End If
            
        End If
        
    End If
    
    If boolSend = True Then
    ' if user wants to send the messsage, ...
    
        .Send
        ' send the message
        
    Else
    ' if user does not want to send the message, ...
    
        .Display
        ' display the message
        
    End If
    
End With

Set objMessage = Nothing
' clear for memory

Set objOutlook = Nothing
' clear for memory

End Sub


