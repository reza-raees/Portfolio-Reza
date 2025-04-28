'Description: This is an automated Workflow which automatically it checks the Query 'SO_USER_REMINDER_NOTIFICATION_IS_EU' and if any row has been updated then it launches the workflow and sent 
'a notification to the related user or role to do the required action

'Note: There is also another workflow in Interface ION which handle the schdualing for automated launching time of the workflow in every hour and here if there is any row then 
'we retrieve the data related to the object and send the notification

Option Strict Off
Imports System
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports System.Collections.Generic
Imports System.Data
Imports System.Data.Common
Imports Microsoft.VisualBasic


Class ActionScript
 Inherits FcProcFuncSetEventWF

Function wf_start() As Long
    Try
    
        PublishDebuggingInfo("SENDING NOTIFICATION  - Process started")
        
        dim gf as GENERALFUNCTIONS = NEW GENERALFUNCTIONS(me)
        Dim dt as datatable = TableLookupEx("SO_USER_REMINDER_NOTIFICATION_IS_EU","TableLookup")
    	if dt.Rows.Count > 0 then
    	    
    	    for each rw as datarow in dt.Rows
    	        Dim sSymbol as string = ""
                Dim objCode as string = ""
                Dim objStatus as String = ""
                Dim sClassName as string = ""
                Dim sDescription as string = ""
                dim sStatusDesc as string = ""
                
    	        Dim sAddress as string
                Dim sAddressType as integer = 1
                if IsBlank(rw("ROLE_CODE")) = 1 then
                    if IsBlank(rw("GROUP_CODE")) = 0 then
                        sAddress = rw("GROUP_CODE")
                        sAddressType = 2
                    end if
                Else
                    sAddress = rw("ROLE_CODE")
                end if 
                
                Dim sSite As String = ""
                
                '===========
                'FORMULA
                '===========
                if IsBlank(rw("OBJ_FORMULA_CODE")) = 0 then
                    
                    sSymbol = "FORMULA"
                    objCode = rw("OBJ_FORMULA_CODE") + "\" + rw("VERSION")
                    objStatus = rw("FORMULA_STATUS")
                    sClassName  = rw("FORMULA_CLASS")
                    sDescription = rw("FORMULA_DESCRIPTION")
                    sStatusDesc = gf.GetStatusDesc("FORMULA", objStatus)
            		sStatusDesc = sStatusDesc & " (" & objStatus & ")"
            		
                    Dim sWorkcenter as Object = ObjProperty("ATTRIBVAL.CONTEXT", "FORMULA" , objCode , "*" , 1)
                    Dim sFacilityDescription As String = ""
                    Dim firstFacility As Object = Nothing
                
                	For x As Integer = 0 To sWorkcenter.Length - 1
                        Dim sWarehouse As Object = ObjProperty("PARENT_CODE", "LOCATION", sWorkcenter(x))
                        Dim sFacility As Object = ObjProperty("PARENT_CODE", "LOCATION", sWarehouse)
                        sFacilityDescription = ObjProperty("DESCRIPTION", "LOCATION", sFacility)
                
                        If firstFacility IsNot Nothing AndAlso sFacility <> firstFacility Then
                            sSite = sSite & " \" & sFacilityDescription
                        ElseIf firstFacility Is Nothing Then
                            firstFacility = sFacility
                            sSite = sFacilityDescription
                        End If
                	Next
                	
                	ObjPropertySet(Now() ,0,"DATE_IND","FORMULA",objCode)
                	
            	else 
        	    '==========
                'ITEM
                '==========
                    sSymbol = "ITEM"
            	    objCode = rw("ITEM_CODE")
                    objStatus = rw("ITEM_STATUS")
                    sClassName = rw("ITEM_CLASS")
                    sDescription = rw("ITEM_DESCRIPTION")
                    sStatusDesc = gf.GetStatusDesc("ITEM", objStatus)
            		sStatusDesc = sStatusDesc & " (" & objStatus & ")"
                    Dim sWarehouse as Object = ObjProperty("ATTRIBVAL.CONTEXT", "ITEM" , objCode , "*" , 1)
                    Dim sFacilityDescription As String = ""
                    Dim firstFacility As Object = Nothing
                
                	For x As Integer = 0 To sWarehouse.Length - 1
                        'Dim sWarehouse As Object = ObjProperty("PARENT_CODE", "LOCATION", sWorkcenter(x))
                        Dim sFacility As Object = ObjProperty("PARENT_CODE", "LOCATION", sWarehouse(x))
                        sFacilityDescription = ObjProperty("DESCRIPTION", "LOCATION", sFacility)
                
                        If firstFacility IsNot Nothing AndAlso sFacility <> firstFacility Then
                            sSite = sSite & " \" & sFacilityDescription
                        ElseIf firstFacility Is Nothing Then
                            firstFacility = sFacility
                            sSite = sFacilityDescription
                        End If
                	Next
                	
                	ObjPropertySet(Now() ,0,"DATE_IND","ITEM",objCode)
                	
                end if   
                
        		Dim sActionDescription as string = rw("ACTION_DESCRIPTION")
            	Dim sWorkflowDesc As String  = rw("WF_DESCRIPTION") 
            	Dim sSourceUser as string = rw("SOURCE_USER_CODE")
            	Dim sWIPCode as string = rw("ACTIONWIP_ID")
            	Dim sActionsetCode as string =rw("ACTIONSET_CODE")
            	
            	Dim wfReason As String
            	if not IsDBNull(rw("WF_REASON_PVALUE")) then wfReason = rw("WF_REASON_PVALUE")
            	
            	Dim sComment as string
            	if not IsDBNull(rw("WIP_COMMENT")) then sComment = rw("WIP_COMMENT")
            	
				'notification function to sent the email by the HTML 'SO_NEW_TASK_APPROVE' template, all the data are mapping into the HTML template which we are designed according to the customer requirment
    	        Notify(sAddress,"SO_NEW_TASK_APPROVE", 1 , sAddressType ,sClassName, objCode ,sDescription,sSourceUser,sWIPCode, sActionsetCode,sWorkflowDesc,sStatusDesc ,Now(),sSite, wfReason ,sActionDescription,sAddress,"",sSymbol,"","")
    	        
    	    next 
    	end if 
        
        PublishDebuggingInfo("SENDING NOTIFICATION  - Process Ended")
        'MessageList("SENDING NOTIFICATION  - Process Ended")
        
        Return 111
    
    Catch ex as exception
        messagelist("Error: "+ ex.message)
        Return 1
        
    end Try
    
End Function

End Class