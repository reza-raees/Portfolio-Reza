' Purpose:  Generating XML file of a Formula and Share it via Interface 

' AUTHOR    :   Reza		September 2024

' NOTE: 	This file includes only the custom-developed logic for Developing script for Publishing the report into the Interface
' 			It does not contain full Interface Configuration or system templates generated and the rest of development is done in the INterface environment.
        
       
'=========================================================================================
'Start _ Changing the structure of the XML according to the Interfac BOD and Publishing 
'=========================================================================================
Dim scurrentTime As String = DateTime.UtcNow.ToString("yyyy-MM-ddTHH:mm:ss.fffffffZ")
Dim xelem As XElement

xelem = XElement.Parse(xmlData)

Dim newXml As XElement =
<SyncDIGIReport xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
	<ApplicationArea>
		<Sender>
			<LogicalID>lid://infor.plmprocess.optiva01</LogicalID>
			<ComponentID>PLM</ComponentID>
			<TaskID>99ace707-daa7-4be6-b8db-fc9dac87d110</TaskID>
		</Sender>
		<CreationDateTime><%= scurrentTime %></CreationDateTime>
		<BODID>01980722-001$BODID$b877257cada44e2fa8e58692917517f3</BODID>
	</ApplicationArea>
	<DataArea>
		<Sync>
			<TenantID>SPA_DEM</TenantID>
			<AccountingEntityID>850_I01</AccountingEntityID>
			<ActionCriteria>
				<ActionExpression actionCode="Change" />
			</ActionCriteria>
		</Sync>
		<DIGIReport>
			<%= xelem.Element("report") %>
		</DIGIReport>
	</DataArea>
</SyncDIGIReport>

xmlData = newXml.ToString()

Dim elem As XElement
Try
	elem = XElement.Parse(xmlData)
Catch ex As Exception
	Throw New Exception("Failed to parse XML data: " & ex.Message)
End Try

Dim theBod As XDocument = New XDocument(elem)
_BODType = "Sync.DIGIReport"  
_ToLogicalId = "lid://infor.plmprocess.optiva01"  
If theBod Is Nothing OrElse theBod.Root Is Nothing Then
	Throw New Exception("theBod is null or empty after creation.")
End If

PublishIONDoc(theBod, "FORMULA" , formulakey , _ToLogicalId)

'messagelist(xmlData)
messagelist("Digitimao Xml file of " + _OBJECTKEY + " Sent to Shared Folder")
PublishDebuggingInfo("Digitimao Xml file of " + _OBJECTKEY + " Sent to Shared Folder")

'=========================================================================================
'End _ Changing the structure of the XML according to the Interfac BOD and Publishing  
'=========================================================================================
 


