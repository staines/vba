Attribute VB_Name = "modCrosstabReport"
Option Compare Database
' updated 2013.06.27
' created by Chris Staines
'
' move controls on a subreport of a crosstab query based on the count of each field,
' thereby allowing for a selective (read:  easily consumable) view of results
'
' usage:
'   1.  create a crosstab query of your liking
'       - enter all column headings in the Column Headings property of Query Properties
'           - this ensures all column headings are displayed, even if data is not provided for each column
'   2.  if necessary, create a query, from the crosstab query, to change counts to x's or another mask
'       - use iif statements, and ensure the name of each field is the same as the original
'   3.  create a query, from the crosstab query, which totals each column as a count
'       - this query will be your strQuery_CrosstabCount
'       - ensure the name of each field is the same as on the original crosstab, and only include fields you would want to hide
'   4.  drag & drop the crosstab query from step 1 (or the query from step 2, if necessary) into a new, blank report
'       - a subreport will be created, which will house your crosstab query
'   5.  include a call to Report_Initiate in the Report_Open subroutine of the resulting subreport (NOT the main report)
'       - hint:  to view the code of a report, select View Code beside the Property Sheet toggle of the Design menu
'       - include an if statement to only perform Report_Initiate when the form's currentview is 6 (report view)

Public Sub Report_Reset(rptCrosstab As Report, intControlMovement As Integer)
' reset the state of each control on the report

Dim objControl As Control
' variable for the control based on the index

Dim intControlIndex As Integer
' variable for the index of the control

Dim intLabelCount As Integer
' variable for the count of labels on the report

intLabelCount% = 0
' set default label count

Dim intTextBoxCount As Integer
' variable for the count of textboxes on the report

intTextBoxCount% = 0
' set default textbox count

For Each objControl In rptCrosstab.Controls
' looping through the controls on the report, ...

    objControl.Visible = True
    ' make the control visible

    Select Case objControl.ControlType
    ' based on the type of control, ...
    
        Case acLabel
        ' if a label, ...

            intLabelCount% = intLabelCount% + 1
            ' increment the label count by 1
            
            If intLabelCount% > 1 Then
            ' if the first label not retained, ...
            
                objControl.Left = intLabelCount% * intControlMovement%
                ' move the label to its default position (its index * movement amount)
                
            End If
            
        Case acTextBox
        ' if a textbox, ...
        
            intTextBoxCount% = intTextBoxCount% + 1
            ' increment the textbox count by 1
            
            If intTextBoxCount% > 1 Then
            ' if the first textbox not retained, ...
            
                objControl.Left = intTextBoxCount% * intControlMovement%
                ' move the textbox to its default position (its index * movement amount)
                
            End If
        
    End Select
    
Next objControl
' continue looping through the controls on the report

End Sub

Private Function Report_ControlIndexByName(rptCrosstab As Report, strControlName As String) As Integer
' obtain a control index using its name

Dim intControl As Integer
' variable for index of current control

For intControl% = 0 To rptCrosstab.Controls.Count - 1
' looping through the count of controls on the report, ...

    If rptCrosstab.Controls(intControl%).Name = strControlName$ Then
    ' if the control's name matches the target name, ...
        
        Report_ControlIndexByName = intControl%
        ' return the control index to the origin function
        
        Exit Function
        ' exit to avoid overlap
        
    End If
    
Next intControl%
' continue looping through the count of controls on the report

End Function

Public Sub Report_Initiate(rptCrosstab As Report, strQuery_CrosstabCount As String, intControlMovement As Integer)
' reset the state of each control and hide/move controls based on their related count

If rptCrosstab.CurrentView = 6 Then Call Report_Reset(rptCrosstab, intControlMovement%)
' reset the report if in report view; it is strongly advised to leave the report as-is if in print view (5)

Dim rsCrosstabCount As Recordset
' variable for recordset of crosstab count query

Set rsCrosstabCount = CurrentDb.OpenRecordset(strQuery_CrosstabCount$, dbOpenDynaset)
' open the crosstab count query

Dim objField As Field
' variable for field in crosstab count query

Dim objControl As Control
' variable for controls on the report

Dim boolControlHidden As Boolean
' variable for if the current control was hidden (determines if subsequent controls should be moved)

Dim intCurrentControlIndex As Integer
' variable for index of current control (if not visible)

Dim intControlIndex As Integer
' variable for index of control after current control (if current not visible)

For Each objField In rsCrosstabCount.Fields
' looping through each field in the crosstab count query, ...

    For Each objControl In rptCrosstab.Controls
    ' looping through each control on the report, ...
    
        boolControlHidden = False
        ' reset variable to avoid moving incorrect control
    
        Select Case objControl.ControlType
        ' based on the type of the control, ...
        
            Case acLabel
            ' if the control is a label (column header), ...
            
                If (objControl.Properties("Caption").Value = objField.Name) And (objField.Value = 0) Then
                ' if the caption is the same as the field name and the field sum is empty, ...
                
                    objControl.Visible = False: boolControlHidden = True
                    ' hide the control
                    
                End If
                
            Case acTextBox
            ' if the control is a textbox (row field of the column), ..

                If (objControl.Name = objField.Name) And (objField.Value = 0) Then
                ' if the report control has the same name as the query field, and the field sum is empty, ...
                
                    objControl.Visible = False: boolControlHidden = True
                    ' hide the control
                
                End If
                
        End Select
        
        If boolControlHidden = True Then
        ' if the current control was hidden, ...
        
            intCurrentControlIndex% = Report_ControlIndexByName(rptCrosstab, objControl.Name)
            ' obtain the index of the current control

            For intControlIndex% = (intCurrentControlIndex%) To rptCrosstab.Controls.Count - 1
            ' for each control from the current control onward, ...
                
                If objControl.ControlType = rptCrosstab.Controls(intControlIndex%).ControlType Then
                ' if the type of the current control is the same as the type of the subsequent control, ...
                
                    rptCrosstab.Controls(intControlIndex%).Left = rptCrosstab.Controls(intControlIndex%).Left - intControlMovement%
                    ' move the subsequent control the control movement amount
                    
                    rptCrosstab.Width = rptCrosstab.Width - intControlMovement%
                    ' reduce the size of the report to match the hiding of controls
                    
                End If
            
            Next intControlIndex%
            ' continue to the next control

        End If
    
    Next objControl
    ' continue looping through each control on the report
    
Next objField
' continue looping through each field in the crosstab count query

End Sub
