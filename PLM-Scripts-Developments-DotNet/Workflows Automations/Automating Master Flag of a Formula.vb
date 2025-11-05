'CREATION A FUNCTION OF AUTOMATING THE MASTER FLAG OF A FORMULA

Function wf_start() As Long
    
Try
    if ObjProperty("STATUS_IND") > 200 then 
        messagelist("Error: The Status of the Formula for the Master Flag Assignment must be equal to or less than 200")
        return 9111
    end if 
    '============
    'Master Flag
    '============
    Dim obj_code = ObjProperty("KEYCODE", "", "")
    Dim mfgItem as Object = ObjProperty("ITEMCODE")
    Dim FoCode as Object = obj_code.Split("-")(0)
	Dim SeqCode as string = obj_code.Substring(obj_code.IndexOf("-") + 1, 3 )	
    Dim currFormulaCode as String = ObjProperty("FORMULA_CODE")
	
    Dim masFlagtable as DataTable = TableLookupEx("SO_ALL_FORMULA_MASFLAG","TableLookup",FoCode)
    ''Assign this formula as master when there is no master flag and set STRT = I01
    if masFlagtable.Rows.count = 0 then 
        ObjPropertySet(1,1,"PRIMARYFORMULAIND", "", "")
	    ObjPropertySet("I01",1,"C_M3STRUCTTYPE", "", "")
	    ObjPropertySet (1,1,"C_FACILITY_MASTER","","")
	    messagelist("Master Flag Enabled")
    else
        'check the FormulaCode of the CurrentFormula with MasterFormula and set the mfgItem to all alternative formula
        for each row as datarow in masFlagtable.Rows
            Dim masFormulaCode as String = row("FORMULA_CODE")
            if currFormulaCode = masFormulaCode Then
                ObjPropertySet(1,1,"PRIMARYFORMULAIND", "", "")
                ObjPropertyset("I01", 1 , "C_M3STRUCTTYPE","","")
                ObjPropertySet (1,1,"C_FACILITY_MASTER","","")
                messagelist("Master Flag Enabled")

                'Update Alternative formulas Mfg. Item with new approved item version
                Dim allFORtb as DataTable = TableLookupEx("SO_FORM_SET_MFGITEM","TableLookup",FoCode, SeqCode)
                for each dr as datarow in allFORtb.Rows
                    Dim rowKeycode as string = dr("FORMULA_CODE") + "\" + dr("VERSION")
                    ObjPropertySet(mfgItem, 1, "ITEMCODE", "FORMULA" , rowKeycode)
                next
                
            else if String.IsNullOrEmpty(row("FORMULA_STATUS")) AndAlso CInt(row("FORMULA_STATUS")) < 201 then
                ObjPropertySet(1,1,"PRIMARYFORMULAIND", "", "")
                ObjPropertyset("I01", 1 , "C_M3STRUCTTYPE","","")
                ObjPropertySet (1,1,"C_FACILITY_MASTER","","")
                
                ObjPropertySet(0,1,"PRIMARYFORMULAIND", "FORMULA", row("FORMULA_CODE") + "\" + row("VERSION"))
                ObjPropertyset(row("FORMULA_CODE").Substring(obj_code.IndexOf("-") + 1, 3), 1 , "C_M3STRUCTTYPE", "FORMULA", row("FORMULA_CODE") + "\" + row("VERSION"))
                ObjPropertySet (0,1,"C_FACILITY_MASTER", "FORMULA", row("FORMULA_CODE") + "\" + row("VERSION"))
                
                messagelist("Master Flag enabled and removed from previous master formula " & row("FORMULA_CODE") + "\" + row("VERSION") + " in status " & row("FORMULA_STATUS"))
                
                'Update Alternative formulas Mfg. Item with new approved item version
                Dim allFORtb as DataTable = TableLookupEx("SO_FORM_SET_MFGITEM","TableLookup",FoCode, SeqCode)
                for each dr as datarow in allFORtb.Rows
                    Dim rowKeycode as string = dr("FORMULA_CODE") + "\" + dr("VERSION")
                    ObjPropertySet(mfgItem, 1, "ITEMCODE", "FORMULA" , rowKeycode)
                next
            else
                messagelist("Error: Master Flag is already Assigned to the formula " + row("FORMULA_CODE") + "\" + row("VERSION") + " with status " & row("FORMULA_STATUS") & " and therefore it cannot be Assigned to this Formula ")
                ObjPropertySet(SeqCode,1,"C_M3STRUCTTYPE", "", "")
                ObjPropertySet (0,1,"C_FACILITY_MASTER","","")
            end If
        next
        
    end If
    return 111

Catch ex as exception
    messagelist("Error: "+ ex.message)
    PublishDebuggingInfo("Error: " + ex.message)
    Return 1
End Try
    
End Function


'Creation of a Function to Check and Validate the Existance of a Master Formula in the Related mfg Item
Function Validate_Master_Formula_Check() as Integer
	Dim retval as Integer = 0
	Dim FoCode as Object = co._OBJECTKEY.Split("-")(0)
	
	Dim masFlagtable as DataTable = co.TableLookupEx("SO_ALL_FORMULA_MASFLAG","TableLookup",FoCode)
	'check the existance of the master in alternatives
	if masFlagtable.Rows.count = 0 then 
		retval = -1
		co.messagelist("{font color='red'}Error: There is no Master Formulas Associated to the mfg Item{/font}")
	end if
	
	Dim FacilityCode as string = co.ObjProperty("C_FACILITY")
	Dim masFacilitytable as DataTable = co.TableLookupEx("SO_FACILITY_MASTER_IR_IT","TableLookup",FoCode,FacilityCode)
	'if there was no formula with the masterfacility in the current facility
	if masFacilitytable.Rows.count = 0 then 
		retval = -1
		co.messagelist("{font color='red'}Error: There is no Facility Master Formulas Associated to the mfg Item{/font}")
	end if
	
	Return retval
End Function


'Related Query used in the script'SO_ALL_FORMULA_MASFLAG'

	SELECT TOP 1 FSITEM.ITEM_CODE , FSITEM.STATUS_IND, FSITEM.FORMULA_ID, FSFORMULA.FORMULA_CODE, FSFORMULA.VERSION, FSFORMULA.STATUS_IND AS FORMULA_STATUS
	FROM FSITEM
	LEFT OUTER JOIN FSFORMULA
	ON FSFORMULA.FORMULA_ID = FSITEM.FORMULA_ID
	WHERE FSFORMULA.FORMULA_CODE LIKE [%1] + '-%'