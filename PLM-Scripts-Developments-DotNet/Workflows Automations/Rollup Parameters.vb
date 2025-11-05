'Function of the Rolling up some parameters 


'Begin  -   Reza.  Rolling up the params SO_APPLICATION_EU and SO_APPLICATION_USA 

Function RollUp_Application_Param()

    Dim sClass As String = co.ObjProperty("CLASS") 
    If sClass = "BULK_FORMULA" Or sClass = "DEV_BULK_FORMULA" Then

        Dim itemIngr As Object = co.ObjProperty("ITEMCODE.INGR", "", "", "*", 1)
        Dim itemCom As Object = co.ObjProperty("COMPONENTIND.INGR", "", "", "*", 1)
        Dim frmEuRollup As String = ""
        Dim frmUsRollup As String = ""
        Dim checkEu As Boolean = False
        Dim checkUs As Boolean = False
        Dim blankEuFound As Boolean = False
        Dim blankUsFound As Boolean = False
        Dim notAllowed As Boolean = False
        Dim euValList As New List(Of String)
        Dim usValList As New List(Of String)

        If itemIngr.Length > 0 Then
            For x As Integer = 0 To itemIngr.Length - 1
                If itemCom(x) = "8" Then
                    Dim itemEu As String = co.ObjProperty("VALUE.TPALL", "ITEM", itemIngr(x), "SO_APPLICATION_EU", "PARAM_CODE")
                    If co.isblank(itemEu) = 0 Then
                        If itemEu <> "1" Then 
                            checkEu = True
                            If Not euValList.Contains(itemEu) Then euValList.Add(itemEu)
                        End If
                    Else
                        blankEuFound = True
                    End If

                    Dim itemUs As String = co.ObjProperty("VALUE.TPALL", "ITEM", itemIngr(x), "SO_APPLICATION_USA", "PARAM_CODE")
                    If co.isblank(itemUs) = 0 Then
                        If itemUs = "NOT ALLOWED" Then
                            notAllowed = True
                        ElseIf itemUs <> "COSMETICS INCL EYES" Then
                            checkUs = True
                            If Not usValList.Contains(itemUs) Then usValList.Add(itemUs)
                        End If
																	   
                    Else
                        blankUsFound = True
                    End If
                End If
            Next

            ' EU Rollup
            If blankEuFound Then
                frmEuRollup = ""
            ElseIf checkEu = False Then
                frmEuRollup = "1" ' All "Alle Bereiche"
            Else
                frmEuRollup = String.Join(";", euValList)
            End If

            ' US Rollup
            If blankUsFound Then
                frmUsRollup = ""
            ElseIf notAllowed = True Then
                frmUsRollup = "NOT ALLOWED"
            ElseIf checkUs = False Then
                frmUsRollup = "COSMETICS INCL EYES"
            Else
                frmUsRollup = String.Join(";", usValList)
            End If

            co.messagelist("EU Form Rollup Result : " & frmEuRollup)
            co.messagelist("USA Form Rollup Result : " & frmUsRollup)

            co.ObjPropertySet(frmEuRollup, 1, "VALUE.TPALL", "", "", "SO_APPLICATION_EU", 2)
            co.ObjPropertySet(frmUsRollup, 1, "VALUE.TPALL", "", "", "SO_APPLICATION_USA", 2)
        End If
    End If
End Function

'End    -   Reza.   Rolling up the params SO_APPLICATION_EU and SO_APPLICATION_USA 


'Calling Function in the script to automate the calculation

 Dim sClass As String = ObjProperty("CLASS","","")
if sClass = "BULK_FORMULA" or sClass = "DEV_BULK_FORMULA" then
	genfun.RollUp_Application_Param() 
end if
	
	
	
	
'RollUp of another two Parameters directly	
	
'==============================================================================================
'Begin  -   Reza.     Rolling up the Notes SO_PEELING & SO_STIR_BOTTLING
'==============================================================================================
Dim formIngr as Object = ObjProperty("FORMULACODE.INGR", "", "", "*", 1)

If formIngr.Length > 0 then
	If sClass = "BULK_FORMULA" or sClass = "DEV_BULK_FORMULA" then
		For x As Integer = 0 to formIngr.Length -1
			If IsBlank(formIngr(x)) = 0 then
				Dim stir_bott As String = ObjProperty("DOCTEXT.DOC.A","FORMULA",formIngr(x),"SO_STIR_BOTTLING",1)
				Dim peel      As String = ObjProperty("DOCTEXT.DOC.A","FORMULA",formIngr(x),"SO_PEELING",1)
				ObjPropertySet(stir_bott,1,"DOCTEXT.DOC.A", "", "", "SO_STIR_BOTTLING", 1)
				ObjPropertySet(peel     ,1,"DOCTEXT.DOC.A", "", "", "SO_PEELING", 1)
			End if
		Next
	End if
	If sClass = "ITEM_PACK" then
		For x As Integer = 0 to formIngr.Length -1
			If IsBlank(formIngr(x)) = 0 AndAlso ObjProperty("CLASS","FORMULA",formIngr(x)) <> "BULK_FORMULA" AndAlso ObjProperty("CLASS","FORMULA",formIngr(x)) <> "DEV_BULK_FORMULA" then
				Dim stir_bott As String = ObjProperty("DOCTEXT.DOC.A","FORMULA",formIngr(x),"SO_STIR_BOTTLING",1)
				Dim peel      As String = ObjProperty("DOCTEXT.DOC.A","FORMULA",formIngr(x),"SO_PEELING",1)
				ObjPropertySet(stir_bott,1,"DOCTEXT.DOC.A", "", "", "SO_STIR_BOTTLING", 1)
				ObjPropertySet(peel     ,1,"DOCTEXT.DOC.A", "", "", "SO_PEELING", 1)
			End if
		Next
	End if
End if
'==============================================================================================
'End    -   Reza.     Rolling up the Notes SO_PEELING & SO_STIR_BOTTLING
'==============================================================================================