'Removing the Alternative Ingredients from the created object 

'Reza         - 17/07/2025  Removing Alternative Ingredient if it is a Save As
	if _MODELOBJECTKey <> "" and _MODELOBJECTKey <> "@DFLT" then 
        Dim ds As DataSet = ObjectDataSet("", "")
        Dim altIngrTableName As String = DataSetTableName("", "", "ALTINGR")
        If ds.Tables.Contains(altIngrTableName) Then
            Dim altIngrTable As DataTable = ds.Tables(altIngrTableName)
             For i As Integer = altIngrTable.Rows.Count - 1 To 0 Step -1 
                altIngrTable.Rows(i).Delete()
            Next
        End If
	end If
	'Reza            01/08/2025  valorizing RD_CECNTER param automatically