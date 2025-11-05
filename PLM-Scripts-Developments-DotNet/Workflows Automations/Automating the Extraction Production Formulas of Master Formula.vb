'Creation of a Function to retrieve Default Productionformulas of Master Formula in "SAVE" script
    Function Defaul_Production_of_Master_Formula(frmCode as String) As String 
        
        Dim defprdformulalist As New List(Of String)
        Dim FoCode as Object = frmCode.ToString().Split("-"c)(0)
		Dim attribArr As String() = co.ObjProperty("ATTRIBCODE.CONTEXT", "FORMULA", frmCode, "*")
		Dim valuesArr As String() = co.ObjProperty("ATTRIBVAL.CONTEXT", "FORMULA", frmCode, "*") 
		dim prdFormCode as string = ""
		For i As Integer = 0 To valuesArr.Length - 1
			If attribArr(i).Equals("MFGLOC") Then
				Dim warehouse_code as String = co.ObjProperty("PARENT_CODE", "LOCATION",valuesArr(i))
				Dim facility_code as string = co.ObjProperty("PARENT_CODE", "LOCATION",warehouse_code)
				Dim facilityDesc As String = co.ObjProperty("DESCRIPTION", "LOCATION", facility_code)
				Dim tb as dataTable = co.TableLookupEx("SO_DEFAULT_PROD_FLAG","TableLookup",facility_code,FoCode)
				if tb.Rows.count > 0 then
					for each rw as DataRow in tb.Rows 
					    dim prdformfac as string = rw("FORMULA_CODE")+"\"+rw("VERSION")+"("+facility_code+")"
					    if not defprdformulalist.Contains(prdformfac) then
					        defprdformulalist.Add(prdformfac)
							if prdFormCode = "" then
								prdFormCode = prdformfac
							else
								prdFormCode = prdFormCode & ";" &
											  prdformfac
							END if
					    end if
					Next
				end if
				
			end if
		Next
        Return prdFormCode
        
    End Function
	
'Creation of a function to Check and Validate the Existance of the Production Flag of a formula 
	Function Validate_Check_Default_Production_Flag() as integer
        Dim retval as Integer = 0
        Dim facility_list as new list(Of String)
        Dim attribArr As String() = co.ObjProperty("ATTRIBCODE.CONTEXT", "FORMULA", co._OBJECTKEY, "*")
		Dim valuesArr As String() = co.ObjProperty("ATTRIBVAL.CONTEXT", "FORMULA", co._OBJECTKEY, "*") 
		Dim FoCode as Object = co._OBJECTKEY.Split("-")(0)
		
        For i As Integer = 0 To valuesArr.Length - 1
            If attribArr(i).Equals("MFGLOC") Then
                
				Dim warehouse_code as String = co.ObjProperty("PARENT_CODE", "LOCATION",valuesArr(i))
				Dim facility_code as string = co.ObjProperty("PARENT_CODE", "LOCATION",warehouse_code)
				Dim facilityDesc As String = co.ObjProperty("DESCRIPTION", "LOCATION", facility_code)
				Dim tb as dataTable = co.TableLookupEx("SO_DEFAULT_PROD_FLAG","TableLookup",facility_code,FoCode)
				
				if tb.Rows.count = 0 then
				    if not facility_list.Contains(facility_code) then
    				    facility_list.add(facility_code)
    				    if co._WIPLINEID <> 1 then
    				        retval = -1
    				        if co._ACTIONCODE <> "SO_SF_RD_MDATA_IR_IT" then
    				            co.messagelist("{font color='red'}Error: There is no Alternative Formula with Enabled Production Flag in the Facility of " + facilityDesc + "{/font}")
    				        end if
    			        else 
    			            co.messagelist("{font color='red'}Warning: There is no Alternative Formula with Enabled Production Flag in the Facility of " + facilityDesc + "{/font}")
    				    end if
				    end if
				end if
				
			end if
        Next

        Return retval
    End Function
	
	
		'Related Query info used in script 'SO_DEFAULT_PROD_FLAG'

		SELECT * FROM FSFORMULA F
		JOIN FSFORMULAATTRIB B ON B.FORMULA_ID = F.FORMULA_ID
		JOIN FSLOCATION L1 ON L1.LOCATION_CODE = B.ATTRIB_VAL
		JOIN FSLOCATION L2 ON L2.LOCATION_CODE = L1.PARENT_CODE
		WHERE FORMULA_CODE LIKE [%2]+'-%'
		AND L2.PARENT_CODE LIKE [%1]
		AND F.C_PRODUCTION_FORMULA = 1
		AND F.STATUS_IND < 800