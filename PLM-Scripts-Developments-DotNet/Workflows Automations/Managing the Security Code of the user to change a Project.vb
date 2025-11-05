

'Creation of a Workflow to manage the Security Code for an specific User 
Class ActionScript
Inherits FcProcFuncSetEventWF

Function wf_start() As Long
Dim statuscode As Object = ObjProperty("STATUSIND.STATUS","","")
Dim iCount As Integer = 0

If statuscode <> "900" Then
    MessageList("The Status of the Project is not CLOSED")
    iCount = 1   
End If

If iCount > 0  Then
    Messagelist("Workflow canceled.")
           Return 9111
End If
'*************************************************************************************************************
' Change Security Object from LOCK to UNLOCK in order to permit modifications  
'*************************************************************************************************************

SetSecurity("", "", 7,7,3)

ObjPropertySet("910", 0, "STATUSIND.STATUS", "", "")
ObjPropertySet("PENDING", 0, "APPROVALCODE.STATUS", "", "")

 Dim wfhmsg As String = "has Started the Workflow " & _ACTIONSETCODE
 Dim cHistory As HISTORY = New HISTORY(Me)
 cHistory.wf_history(_OBJECTKEY, _OBJECTSYMBOL, _STARTUSER, "HISTORY", wfhmsg)'

 MessageList("Project  is now ready for change; After your change, you have to complete in the pending task list" )

 Return 1

End Function

Function wf_complete() As Long
    
'*************************************************************************************************************
' Change Security Object to  re-=LOCK  
'*************************************************************************************************************
    SetSecurity("", "", 3, 3 ,3)
    
    ObjPropertySet("900", 0, "STATUSIND.STATUS", "", "")
    ObjPropertySet("APPROVED", 0, "APPROVALCODE.STATUS", "", "")
    Dim wfhmsg As String = "has completed the Workflow " & _ACTIONSETCODE
    Dim cHistory As HISTORY = New HISTORY(Me)
    cHistory.wf_history(_OBJECTKEY, _OBJECTSYMBOL, _SOURCEUSER, "HISTORY", wfhmsg)
    
    Messagelist("You have completed this Project change. ")
Return 111
End Function


End Class