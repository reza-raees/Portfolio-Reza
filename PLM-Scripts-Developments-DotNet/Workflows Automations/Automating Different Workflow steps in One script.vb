'===================================================================================================
'Begin :  Reza - Automating Different Workflows Steps at the start point according to the Wf's Name   
'===================================================================================================
if _ACTIONSETCODE = "SO_SF_APPROVAL_1_IS_US" then 
        dim oItems as object = ObjProperty("ITEMCODE.INGR.A", "", "", "*", 1)
        
        'check the status of all itemING to be 200 
        Dim allstatusabove200 as Boolean = True
        For x as Integer = 0 to oItems.Length -1 
            Dim itemINGstatus as Object = ObjProperty("STATUSIND.STATUS","ITEM",oItems(x))
            if itemINGstatus < 200 then
                allstatusabove200 = False
                Exit For
            end If
        next
        
        If allstatusabove200 = True then

            Dim EqualWarehouse as Boolean = False ' Assume True initially
            Dim sfFormulaContext As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)
            For f as Integer = 0 to sfFormulaContext.Length -1
                Dim warehouseSFformula as String = ObjProperty("PARENT_CODE", "LOCATION", sfFormulaContext(f))
                'messagelist("the workcenter is: ",sfFormulaContext(f),"and its warehouse is: ",warehouseSFformula)
                    
                For r as Integer = 0 to oItems.Length - 1
                    Dim itemINGContext As Object = ObjProperty("ATTRIBVAL.CONTEXT", "ITEM", oItems(r), "*", 1)
                   
                    For m as Integer = 0 to itemINGContext.Length - 1
                        'messagelist("item ingeridients: ", oItems(r) ," and its warehouse is: ",itemINGContext(m))
                        if itemINGContext(m) = warehouseSFformula then
                            EqualWarehouse = True
                            Exit For ' Exit the loop since we found a match
                        else
                            EqualWarehouse = False
                        end if
                    next
                    if EqualWarehouse = False Then
                        Exit For
                    end if
                next
                if EqualWarehouse = False Then
                    Exit For
                end If
            next
            
            If EqualWarehouse = True then
                WSkip("SO_FRM_SF_APPR_REJ_IS_US")
                
                if not WipParamGet("CREATE_FG") is Nothing AndAlso WipParamGet("CREATE_FG") = 1 then
                    'generating new code for new FG Formula   
                    Dim sDescription as string = Objproperty("DESCRIPTION")
                    Dim sUom as String = Objproperty("UOM_CODE")
                    Dim sObjectCode as Object = GenerateCodeNumber("SO NEW AUTOCREATION FP IS_US")   'sObjectCodesFGformulaCode
                    Dim sSourceUser as String = _SOURCEUSER
                    Dim cFGformula As SO_FORMULACREATE  = New SO_FORMULACREATE(Me)
                    cFGformula.cr_FGformula_IS_US(sObjectCode , sDescription, sUom , sSourceUser)
                    MessageList("New Related Finished Good Bom is:", sObjectCode)
                    ObjPropertySet("PENDING", 0, "APPROVALCODE.STATUS", "FORMULA", sObjectCode)
                    
                    'generating new mfgItem for the new generated FG formula
                    Dim FoCode as String = sObjectCode.Split("-")(0)
                    Dim sItemCode As String = FoCode + "-" + ObjProperty("VERSION","","")
                    Dim cMfgItem As SO_MFGITEMCREATE = New SO_MFGITEMCREATE(Me)
                    cMfgItem.cr_mfgItem_IS_US(sItemCode, sDescription, sUom, sSourceUser)
                    ObjPropertySet(sItemCode , 1 , "ITEM_CODE" , "FORMULA" , sObjectCode)
                    
                    'Add the mfg of sf to the INGR of the fg
                    Dim sfMfgItem as String = Objproperty("ITEM_CODE","","")
                    Dim ds As DataSet = ObjectDataSet("FORMULA",sObjectCode)
                    Dim IngrTableName As String = DataSetTableName("FORMULA",sObjectCode, "INGR")
                    Dim IngrTable As DataTable = ds.Tables(IngrTableName)
                    Dim newIngrRow As DataRow = GetNewRow("FORMULA", sObjectCode ,"INGR")
                    newIngrRow("ITEM_CODE") = sfMfgItem
                    'newIngrRow("ITEM_FORMULA_ID") = Objproperty("FORMULA_CODE","","") + "\" + ObjProperty("VERSION","","")
                    newIngrRow("QUANTITY") = 1000.0
                    newIngrRow("ACTIVE_QUANTITY") = 1000.0
                    RowUpdate("FORMULA", sObjectCode , "INGR", newIngrRow)
                    CommitNewRow("FORMULA", sObjectCode , "INGR", newIngrRow)
                    
                    'add context of SF to the FG context
                    Dim sfContext as Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)
                    For j As Integer = 0 To sfContext.Length - 1
                        Dim addFGworkcenter As Long = AddContextAttrib("FORMULA", sObjectCode, "MFGLOC", sfContext(j))
                    Next
                    
                    
                    'Copy all validated parameters from SF to FG
                    Dim validatedParam As New List(Of String) From {"SHELF_LIFE","PRODUCT_FORMAT","PRODUCT_MINOR_CODE","RD_CENTRE","PAESE_ULT_LAVORAZ","FLASH_POINT_TEMP","PHYSICAL_STATE"}
                    For Each param As String In validatedParam
                        
                        Dim paramValue as string = ObjProperty("VALUE.TPALL", "", "", param , "PARAM_CODE") 
                        ObjPropertySet(paramValue , 0 , "VALUE.TPALL" , "FORMULA" , sObjectCode , param , "PARAM_CODE" )
                    Next
                    
                    Dim noteparams As New List(Of String) From {"PHYSICAL_DESCRIPTION","STORAGE_HANDLING","FLAVOUR","COLOUR","ODOUR"}
                    For Each note As String In noteparams
                        Dim notesValue as string = ObjProperty("DOCTEXT.DOC.A", "", "", note ,1 )
                        ObjPropertySet(notesValue,0,"DOCTEXT.DOC.A", "", sObjectCode, note ,1)
                    Next
                    StartWorkflow("SO_FP_COMPLETE_PK_IS_US", "FORMULA" , sObjectCode)

                End if
                
           end if

        end if 
	end if
	
	'Other workflows
    if _ACTIONSETCODE = "SO_INTERM_SF_APPROVAL_IS_US" or _ACTIONSETCODE = "SO_ALTER_FRM_APPROVAL_IS_US"  then 
        
         dim oItems as object = ObjProperty("ITEMCODE.INGR.A", "", "", "*", 1)
        
        'check the status of all itemING to be 200 
        Dim allstatusabove200 as Boolean = True
        For x as Integer = 0 to oItems.Length -1 
            Dim itemINGstatus as Object = ObjProperty("STATUSIND.STATUS","ITEM",oItems(x))
            if itemINGstatus < 200 then
                allstatusabove200 = False
                Exit For
            end If
        next
        
        If allstatusabove200 = True then

            Dim EqualWarehouse as Boolean = False ' Assume True initially
            Dim sfFormulaContext As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)
            For f as Integer = 0 to sfFormulaContext.Length -1
                Dim warehouseSFformula as String = ObjProperty("PARENT_CODE", "LOCATION", sfFormulaContext(f))
                'messagelist("the workcenter is: ",sfFormulaContext(f),"and its warehouse is: ",warehouseSFformula)
                    
                For r as Integer = 0 to oItems.Length - 1
                    Dim itemINGContext As Object = ObjProperty("ATTRIBVAL.CONTEXT", "ITEM", oItems(r), "*", 1)
                   
                    For m as Integer = 0 to itemINGContext.Length - 1
                        'messagelist("item ingeridients: ", oItems(r) ," and its warehouse is: ",itemINGContext(m))
                        if itemINGContext(m) = warehouseSFformula then
                            EqualWarehouse = True
                            Exit For ' Exit the loop since we found a match
                        else
                            EqualWarehouse = False
                        end if
                    next
                    if EqualWarehouse = False Then
                        Exit For
                    end if
                next
                if EqualWarehouse = False Then
                    Exit For
                end If
            next
            
            If EqualWarehouse = True then
                if _ACTIONSETCODE = "SO_SF_APPROVAL_2_IS_US" then
                    WSkip("SO_FRM_SF_APPR_REJ_IS_US")
                    WSkip("SO_FRM_SF_APPR_REJ_IS_US")
                    ObjPropertySet("210", 0, "STATUSIND.STATUS", "", "")
                elseif _ACTIONSETCODE = "SO_INTERM_SF_APPROVAL_IS_US" then 
                    WSkip("SO_FRM_INT_SF_APPR_REJ_IS_US")
                    ObjPropertySet("110", 0, "STATUSIND.STATUS", "", "")
                elseif _ACTIONSETCODE="SO_FP_APPROVAL_2_IS_US" then
                    WSkip("SO_FRM_APPR_REJ_IS_US")
                    WSkip("SO_FRM_APPR_REJ_IS_US")
                    ObjPropertySet("210", 0, "STATUSIND.STATUS", "", "")
                elseif _ACTIONSETCODE = "SO_ALTER_FRM_APPROVAL_IS_US" then 
                    WSkip("SO_ALT_FRM_APPR_REJ_IS_US")
                    WSkip("SO_ALT_FRM_APPR_REJ_IS_US")
                    ObjPropertySet("110", 0, "STATUSIND.STATUS", "", "")
                end if
            end if
        end if
    end if
	
	'Other workflows 
    if _ACTIONSETCODE = "SO_FRM_MAT_CHANGE_APPR_IS_US" then 
        dim iUser as string = _STARTUSER
    	dim currentRole as string = "RL_RD_IS_US"
    	dim tb as dataTable = TableLookupEx("USERINROLE" ,"TableLookup", iUser , currentRole)
    	for each dr as datarow in tb.Rows
    	    messagelist(dr("NUM_IN_ROLE"))
    		if dr("NUM_IN_ROLE") > 0 then
        	    ObjPropertySet("762", 0, "STATUSIND.STATUS", "", "")
            else
                Wskip("SO_SITE_RECOGNITION_IS_US")
                Wskip("SO_FRM_MAT_CHANGE_APPR_IS_US")
                Wskip("SO_FRM_MAT_CHANGE_APPR_IS_US")
                Wskip("SO_FRM_MAT_CHANGE_APPR_IS_US")
                WIPInfoSet("GROUP_CODE","", 8)
                WIPInfoSet("ROLE_CODE","RL_RD_IS_US", 8)
                'WIPInfoSet("DESCRIPTION","R&D Review", 7)
                ObjPropertySet("761", 0, "STATUSIND.STATUS", "", "")
        	end if  
    	next
    end if 
	
'===================================================================================================
'End :  Reza - Automating Different Workflows Steps at the start point according to the Wf's Name   
'===================================================================================================