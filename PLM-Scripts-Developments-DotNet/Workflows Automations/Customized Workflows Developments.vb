'Description: Customised Updates and Developments inside different Workflows

'Note: these are all the new developed Parts added to the Workflows and not privided the whole script of the workflows


'===================================================================================================
'Start : Reza - Master Flag Assigned to the Formula which first Approved 04/01/2024
'===================================================================================================
Dim mfgItem as Object = ObjProperty("ITEMCODE")
Dim FoCode as Object = obj_code.Split("-")(0)
Dim SeqCode as string = obj_code.Substring(obj_code.IndexOf("-") + 1, 3 )	
Dim currFormulaCode as String = ObjProperty("FORMULA_CODE","","") ' getting the 8characters + Seq code

'***Start: Reza - AVOIDING ASSIGNING MASTERFLAG FOR FORMULAS WHICH EXISTS IN GLOBAL LAB 09/07/2024 ***
Dim globaltable as DataTable = TableLookupEx("SO_MIGRATION_GLOBAL_TO_IS_EU","TableLookup",FoCode)
if  globaltable.rows.count = 0 then
		
		
'***End: Reza - AVOIDING ASSIGNING MASTERFLAG FOR FORMULAS WHICH EXISTS IN GLOBAL LAB 09/07/2024 ***

	Dim masFlagtable as DataTable = TableLookupEx("SO_ALL_FORMULA_MASFLAG","TableLookup",FoCode)
	'Assign this formula as master when there is no master flag and set STRT = I01
	if masFlagtable.Rows.count = 0 then 

		ObjPropertySet(1,1,"PRIMARYFORMULAIND", "", "")
		ObjPropertySet("I01",1,"C_M3STRUCTTYPE", "", "")

	else
		
		'check the FormulaCode of the CurrentFormula with MasterFormula and set the mfgItem to all alternative formula
		for each row as datarow in masFlagtable.Rows
			Dim masFormulaCode as String = row("FORMULA_CODE")
			if currFormulaCode = masFormulaCode Then
				'enable the masterflag and structure type
				ObjPropertySet(1,1,"PRIMARYFORMULAIND", "", "")
				objpropertyset("I01", 1 , "C_M3STRUCTTYPE","","")
				
				Dim allFORtb as DataTable = TableLookupEx("SO_FORM_SET_MFGITEM","TableLookup",FoCode, SeqCode)
				for each dr as datarow in allFORtb.Rows
					Dim rowKeycode as string = dr("FORMULA_CODE") + "\" + dr("VERSION")
					ObjPropertySet(mfgItem, 1, "ITEMCODE", "FORMULA" , rowKeycode)
				next
				
			else
				
				ObjPropertySet(SeqCode,1,"C_M3STRUCTTYPE", "", "")
			end If
		next
		
	end If
	
end If												
'===================================================================================================
'End : Reza - Master Flag Assigned to the Formula which first Approved 04/01/2024
'===================================================================================================

'===================================================================================================
'Start : Reza - MasterFacility Assigned if its the first in facilityset 08/02/2024
'===================================================================================================
	'check the master facility in the all related formula
	Dim FacilityCode as string = ObjProperty("C_FACILITY","","")
	Dim masFacilitytable as DataTable = TableLookupEx("SO_FACILITY_IS_EU","TableLookup",FoCode,FacilityCode)
	'if there was no formula with the masterfacility in the current facility
	if masFacilitytable.Rows.count = 0 then 
		'set the master facility flag enable for current facility
		objpropertyset (1,1,"C_FACILITY_MASTER","","")
		objpropertyset("I01", 1 , "C_M3STRUCTTYPE","","")
	else
		'there is already a masterfacility then check the version
		for each rw as datarow in masFacilitytable.Rows
			Dim masFacilityForCode as String = rw("FORMULA_CODE")
			if currFormulaCode = masFacilityForCode Then
				ObjPropertySet(1,1,"C_FACILITY_MASTER", "", "")
				objpropertyset("I01", 1 , "C_M3STRUCTTYPE","","")
			else
				ObjPropertySet(0,1,"C_FACILITY_MASTER", "", "")
				ObjPropertySet(SeqCode,1,"C_M3STRUCTTYPE", "", "")
			end if
		next
	end if
'===================================================================================================
'End : Reza - MasterFacility Assigned if its the first in facilityset 08/02/2024
'===================================================================================================


	
'===================================================================================================
'Begin: Reza - Handling MfgItme's status for the Formula product change Wf 13/11/2024
'===================================================================================================
If newstat = 600 or newstat = 700 then	
	Dim mfgItem as Object = ObjProperty("ITEMCODE")
	Dim cformula as string = objproperty("FORMULA_CODE","","")
	
	
	if _ACTIONSETCODE = "SO_FP_MAT_CHANGE_APPR_IS_EU" or _ACTIONSETCODE = "SO_SF_MAT_CHANGE_APPR_IS_EU" then
	
		dim appItemtable as DataTable = TableLookupEx("SO_APPR_MFGITEM_IS_EU","TableLookup", mfgitem)
		
		if appItemtable.Rows.Count = 0 then
			
			if (startM3Interface <=0) then
				
				ObjPropertySet(newstat, 0, "STATUSIND.STATUS", "ITEM", mfgitem)
				ObjPropertySet("APPROVED", 0, "APPROVALCODE.STATUS", "ITEM", mfgitem)
				
			Else
				
				ObjPropertySet(700, 0, "STATUSIND.STATUS", "ITEM", mfgitem)
				ObjPropertySet("APPROVED", 0, "APPROVALCODE.STATUS", "ITEM", mfgitem)
				
			end If	
			
			updateItemHistory = 1
			
		end If
	end if 
	
	'===================================================================================================
	'End: Reza - Handling MfgItme's status for the Formula product change Wf 13/11/2024
	'===================================================================================================
	
	'===================================================================================================
	'Begin: Reza - Adding/removing MfgItme's Warehouse according to the Formula's Workcenter 10/01/2024
	'===================================================================================================
	'Adding new wareHouses for mfgItem according to the related formula
	Dim itemContext As Object = ObjProperty("ATTRIBVAL.CONTEXT", "ITEM", mfgItem, "*", 1)
	Dim parentTable As DataTable = TableLookupEx("SO_FORMULA_WORKCENTER_IS_EU", "TableLookup", mfgItem, cformula)

	For Each parentRow As DataRow In parentTable.Rows   
		Dim parentCode As String = parentRow("PARENT_CODE").ToString()
		Dim matchFound As Boolean = False
	
		For y As Integer = 0 To itemContext.Length - 1
			If itemContext(y).ToString() = parentCode Then
				matchFound = True
				Exit For ' exit the inner loop once a match is found
			End If
		Next
	
		If Not matchFound Then
			' No match was found for parentRow("PARENT_CODE") in itemContext
			Dim wrHouseSell As Long = AddContextAttrib("ITEM", mfgItem, "SELLOC", parentCode)
		End If
	 Next
	 
end if

'===================================================================================================
'End :  Reza - Adding/removing MfgItme's Warehouse according to the Formula's Workcenter 09/01/2024
'===================================================================================================

'===================================================================================================
'Begin :  Reza - cancelling the wf of the old version formula while the new one is approving   
'===================================================================================================
'cancel other runnnig wf in WIP
if (oldStatus < 800 and oldStatus > 700) then
	Dim dts as DataTable = TableLookupEx("SO_CANCEL_WF_IS_EU","TABLELOOKUP",oldkey)
	if dts.Rows.Count > 0 then
		For each row as DataRow in dts.Rows 
			dim oldwipId as String = row("ACTIONWIP_ID")
			messagelist("The Workflow ID "+ oldwipId + " Related to the " + oldkey + " Cancelled")
			ObjMethod("ACTIONWIP" , oldwipId, "ObjectMethod", "RECALL", "")
		next
	end if 
end if 
'===================================================================================================
'End :  Reza - cancelling the wf of the old version formula while the new one is approving   
'===================================================================================================