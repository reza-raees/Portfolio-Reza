
'Creation of a Function to Send an Automatic email when the Role is Skipped in the Workflow

Function Send_emailTo_Skipped_roles(byval sRoleOrGroup as string, byval stepDesc as String)
    
    dim gf as GENERALFUNCTIONS = New GENERALFUNCTIONS(co)
    
    'sending email to the skiipped steps
	
    Dim frmStatus as string = co.objproperty("STATUSIND.STATUS","","")
    Dim sClassName as string = co.ObjProperty("CLASS","","")
    Dim sDescription as string = co.ObjProperty("DESCRIPTION","","")
    dim sStatusDesc as string = gf.GetStatusDesc(co._OBJECTSYMBOL, frmStatus)
		sStatusDesc = sStatusDesc & " (" & frmStatus & ")"
	Dim sWorkflowDesc As String  = co.ObjProperty("DESCRIPTION", "ACTIONSET", co._ACTIONSETCODE) 
	Dim sActionName As String = co.ObjProperty("DESCRIPTION","ACTION",co._ACTIONCODE)
	Dim sWorkcenter as Object = co.ObjProperty("ATTRIBVAL.CONTEXT", "" , "" , "*" , 1)
    Dim sSite As String = ""
    Dim sFacilityDescription As String = ""
    Dim firstFacility As Object = Nothing
    Dim facilities as New List(of String)
    
	For x As Integer = 0 To sWorkcenter.Length - 1
        Dim sWarehouse As Object = co.ObjProperty("PARENT_CODE", "LOCATION", sWorkcenter(x))
        Dim sFacility As Object = co.ObjProperty("PARENT_CODE", "LOCATION", sWarehouse)
        sFacilityDescription = co.ObjProperty("DESCRIPTION", "LOCATION", sFacility)

        If firstFacility IsNot Nothing AndAlso sFacility <> firstFacility Then
            sSite = sSite & " \" & sFacilityDescription
        ElseIf firstFacility Is Nothing Then
            firstFacility = sFacility
            sSite = sFacilityDescription
        End If
        
        If not facilities.Contains(sFacility) Then facilities.Add(sFacility)
	Next

	Dim sIndicator as integer = 1 
	
	if co._ACTIONSETCODE = "SO_SF_APPROVAL_1_IR_IT" or co._ACTIONSETCODE = "SO_FP_APPROVAL_1_IR_IT" or co._ACTIONSETCODE = "SO_ALTER_FP_APPROVAL_IR_IT" then
	    
	    if sRoleOrGroup = "RL_FGRA" or sRoleOrGroup = "RL_QA_LAB" then 
	        co.Notify(sRoleOrGroup ,"SO_TASK_SKIP", 1 , sIndicator ,sClassName, co._OBJECTKEY ,sDescription,co._SOURCEUSER,co._WIPID, co._ACTIONSETCODE,sWorkflowDesc,sStatusDesc ,Now(),sSite,stepDesc,co._TASKROLE,"",co._COMMENT,co._OBJECTSYMBOL)
	    end if
	    
        for each facil as string in facilities
            sRoleOrGroup = "QA_"&facil
            sIndicator = 2
            co.Notify(sRoleOrGroup ,"SO_TASK_SKIP", 1 , sIndicator ,sClassName, co._OBJECTKEY ,sDescription,co._SOURCEUSER,co._WIPID, co._ACTIONSETCODE,sWorkflowDesc,sStatusDesc ,Now(),sSite,stepDesc,co._TASKROLE,"",co._COMMENT,co._OBJECTSYMBOL)
        next
        
    ElseIf co._ACTIONSETCODE = "SO_FRM_CHANGE_IR_IT" then 
        for each facil as string in facilities
            sRoleOrGroup = "OPS_"&facil
            sIndicator = 2
            co.Notify(sRoleOrGroup ,"SO_TASK_SKIP", 1 , sIndicator ,sClassName, co._OBJECTKEY ,sDescription,co._SOURCEUSER,co._WIPID, co._ACTIONSETCODE,sWorkflowDesc,sStatusDesc ,Now(),sSite,stepDesc,co._TASKROLE,"",co._COMMENT,co._OBJECTSYMBOL)
        next
        
    Else
        co.Notify(sRoleOrGroup ,"SO_TASK_SKIP", 1 , sIndicator ,sClassName, co._OBJECTKEY ,sDescription,co._SOURCEUSER,co._WIPID, co._ACTIONSETCODE,sWorkflowDesc,sStatusDesc ,Now(),sSite,stepDesc,co._TASKROLE,"",co._COMMENT,co._OBJECTSYMBOL)
            
	End if
    
    
End Function