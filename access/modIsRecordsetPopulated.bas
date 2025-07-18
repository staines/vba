Attribute VB_Name = "modIsRecordsetPopulated"
Option Compare Database
' updated 2014.05.09
' created by Chris Staines

Function IsRecordsetPopulated(strTableOrQuery As String) As Boolean
' return whether a recordset has contents

Dim rsRecordset As Recordset
' variable for recordset

Set rsRecordset = CurrentDb.OpenRecordset(strTableOrQuery$, dbOpenDynaset)
' open the recordset

If Not (rsRecordset.EOF And rsRecordset.BOF) Then IsRecordsetPopulated = True
' if not at the beginning and end of the recordset (the recordset is not empty), return result to origin function

rsRecordset.Close
' close the recordset

Set rsRecordset = Nothing
' clear the recordset

End Function
