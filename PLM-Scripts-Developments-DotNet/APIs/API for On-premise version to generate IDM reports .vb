' Workflow Action: SO_IDM_REPORT_CALLING_API  

' Purpose:  Generating IDM Report using API (for generating report while IDM.INTEGRATION = 0) 

' AUTHOR    :   Sinfo One
' CREATED   :   04/06/2025
' REVISIONS :   WHO         DATE    	CHANGE  REASON
'               -------     ----------  -----------------------------------------------------------
'               Reza        04/06/2025  Project Sample Report

Option Strict Off
Imports System
Imports System.Collections.Generic
Imports System.Data
Imports System.Diagnostics
Imports System.Globalization
Imports System.IO
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Web
imports System.XML
Imports System.Linq
Imports System.Xml.Linq
Imports System.Threading


Class ActionScript
 Inherits FcProcFuncSetEventWF

Dim cNotification As NOTIFICATION = New NOTIFICATION(Me)
Dim cWfFunctions As WORKFLOWFUNCTIONS = New WORKFLOWFUNCTIONS(Me)
Dim cHistory As HISTORY = New HISTORY(Me)
dim SO_gf as SO_GENERALFUNCTIONS = NEW SO_GENERALFUNCTIONS(me)

Function wf_start() As Long

    PublishDebuggingInfo("Wf Generating IDM Begin")
    
    Try
        
    	'Read values from profile
        Dim p_authTokenEnpoint()        As String   =   GetProfileValue("IDM.TOKENSERVER")
    	Dim p_generateReportEnpoint()   As String   =   GetProfileValue("SO.IDM.API.REPORTGENERATE.URL")
        Dim p_client_id()               As String   =   GetProfileValue("IDM.CLIENT_ID")
        Dim p_client_secret()           As String   =   GetProfileValue("IDM.CLIENT_SECRET")
        Dim p_grant_type()              As String   =   GetProfileValue("IDM.GRANT_TYPE")
        Dim p_username()                As String   =   GetProfileValue("IDM.USERNAME")
        Dim p_password()                As String   =   GetProfileValue("IDM.PASSWORD")
        Dim p_archive_enabled()         As String   =   GetProfileValue("IDM.ARCHIVE_ENABLED")
        Dim p_launchReportEnpoint()     As String   =   GetProfileValue("SO.IDM.API.REPORTLAUNCH.URL")
        
        'Api URL
        Dim authTokenEnpoint            As String   =   p_authTokenEnpoint(0)
        Dim generateReportEnpoint       As String   =   p_generateReportEnpoint(0)
        Dim launchReportEnpoint         As String   =   p_launchReportEnpoint(0)
    	
    	'Authorization parameters
        Dim client_id as String     =   p_client_id(0)
        Dim client_secret as String =   p_client_secret(0)
        Dim grant_type as String    =   p_grant_type(0)
        Dim username as String      =   p_username(0)
        Dim password as String      =   p_password(0)
        
        Dim formContent as FormUrlEncodedContent
        Dim response as HttpResponseMessage
        Dim responseContent as HttpContent
        Dim responseString as String
        
        'Start Calling the API
        
        Using client As HttpClient = New HttpClient()
            
            client.DefaultRequestHeaders.Add("Accept", "application/json;charset=utf-8")
            
            '============================================
            'Get token API
            '============================================
            PublishDebuggingInfo("Get access token")
            Dim token = ""
            
            'Create "application/x-www-form-urlencoded" request content
            Dim params = New Dictionary(Of String, String)
            params.Add("client_id", client_id)
            params.Add("client_secret", client_secret)
            params.Add("grant_type", grant_type)
            params.Add("username", username)
            params.Add("password", password)
            formContent = new FormUrlEncodedContent(params)
            
            'Call API
            response = client.PostAsync(authTokenEnpoint, formContent).GetAwaiter().GetResult()
            responseContent = response.Content
            responseString = responseContent.ReadAsStringAsync().Result
            
            if response.IsSuccessStatusCode then
                PublishDebuggingInfo("Authenticated")
                'MessageList("Response OK: " + responseString)
                
                token = SO_gf.getValueJSON("access_token", responseString)
                            
            else
                
                PublishDebuggingInfo("Get access token failed: " + responseString)
                'MessageList("Response KO: " + responseString)
                
                Dim error_code as String = SO_gf.getValueJSON("error", responseString)
                Dim error_description as String = SO_gf.getValueJSON("error_description", responseString)
                
                Throw new exception("Authentication failed to Salesforce API. Error code: [" + error_code + "] Description: [" + error_description + "]")
            end if
            
           if isblank(token) = 1 then
               Throw new exception("Unale to get access token for Salesforce API")
           end if
           
           client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token)
           
           'Add authorization token
           generateReportEnpoint = generateReportEnpoint + "?" + "Authorization=" + System.Uri.EscapeDataString("Bearer " + token)
           
           'Get Report XML Data
           Dim reportfilename as String = WIPParamGet("FILE_NAME")
           Dim reportxmldata as String = WIPParamGet("XML_REPORT_DATA")
           
           'Create the IDM attribute for storing the item
           Dim idmattribxml As XElement = _
                <item xmlns="http://infor.com/daf">
                    <entityName>OPTIVA_REPORT_OUTPUT</entityName>
                    <attrs>
                        <attr>
                            <name>OPTIVA_SYMBOL</name>
                            <type>0</type>
                            <qual>OPTIVA_SYMBOL</qual>
                            <value><%= _OBJECTSYMBOL %></value>
                        </attr>
                        <attr>
                            <name>OPTIVA_SYMBOL_ID</name>
                            <type>0</type>
                            <qual>OPTIVA_SYMBOL_ID</qual>
                            <value><%= ObjProperty("PROJECT_CODE") %></value>
                        </attr>
                        <attr>
                        	<name>OPTIVA_SYMBOL_VERSION</name>
                        	<type>0</type>
                        	<qual>OPTIVA_SYMBOL_VERSION</qual>
                        	<value><%= ObjProperty("VERSION") %></value>
                        </attr>
                        <attr>
                            <name>OPTIVA_ARCHIVE</name>
                            <type>0</type>
                            <qual>OPTIVA_ARCHIVE</qual>
                            <value>False</value>
                        </attr>
                        <attr>
                            <name>OPTIVA_DOC_TITLE</name>
                            <type>0</type>
                            <qual>OPTIVA_DOC_TITLE</qual>
                            <value><%= reportfilename %></value>
                        </attr>
                    </attrs>
                </item>
            
            'Converting the report xml and the IDM attributes into base64
            Dim idmattribxmlbase64 As String = Convert.ToBase64String(Encoding.UTF8.GetBytes(idmattribxml.ToString()))
            Dim reportxmldatabase64 As String = Convert.ToBase64String(Encoding.UTF8.GetBytes(reportxmldata))
            
            'MessageList("sec xml report 64base: ",reportxmldatabase64)
            'MessageList("sec filename 64base: ",reportfilename)
            
            ' Create JSON request body
            Dim jsonContent As String = _
            "{" + _
                """input"": [" + _
                    "{" + _
                    """type"": ""generate""," + _
                    """data"": {" + _
                        """type"": ""data""," + _
                        """filename"": ""test.xml""," + _
                        """base64"": """+ reportxmldatabase64 + """ "  + _
                                 "}," + _
                    """template"": {" + _
                        """type"": ""xquery""," + _
                        """xquery"": ""/OPTIVA_REPORT_TEMPLATES[@MDS_TemplateName = \""Sample IDM Template\""]"" " + _
                            "}," + _
                        """filename"": """+ reportfilename + """ " + _
                        "}" + _
                    "]," + _
                """targets"": [" + _
                        "{" + _
                    """type"": ""item""," + _
                    """itemdatafile"": {" + _
                        """type"": ""data""," + _
                        """base64"": """+ idmattribxmlbase64 + """ "  + _
                            "}" + _
                        "}" + _
                    "]," + _
                """batchId"": ""string""," + _
                """largeJob"": true" + _
                "}"

            PublishDebuggingInfo("JSON request body: " + jsonContent)
            'MessageList("JSON request body: " + jsonContent)
            
            'Call the  API
            Dim buffer = Encoding.UTF8.GetBytes(jsonContent)
            Dim bytes = New ByteArrayContent(buffer)
            bytes.Headers.ContentType = New Headers.MediaTypeHeaderValue("application/json")
            
            'Call API
            Dim jobId As String
            response = client.PostAsync(generateReportEnpoint, bytes).GetAwaiter().GetResult()
            responseContent = response.Content
            responseString = responseContent.ReadAsStringAsync().Result
            
            Dim message As String = ""
            If responseString.Contains("""success"":true") Then
                
                jobId = SO_gf.getValueJSON("jobId", responseString)
                PublishDebuggingInfo("Success. Job ID: " & jobId)
                'MessageList("Success. Job ID: " & jobId)
                
                'Call API for GetRequest to get the report
                Dim sStatus As String = ""
                'MessageList("launch url:",launchReportEnpoint & "/" & jobId)
                for i as integer = 1 to 120
                    response = client.GetAsync(launchReportEnpoint & "/" & jobId).GetAwaiter().GetResult()
                    responseContent = response.Content
                    responseString = responseContent.ReadAsStringAsync().Result
                    Thread.Sleep(1000)
                    sStatus = SO_gf.getValueJSON("status", responseString)
                    if sStatus = "ok" then
                        'MessageList("status is: " & sStatus)
                        exit for
                    end if
                    i = i+1
                next
                
                dim oidm as IDM_EPPLUS = New IDM_EPPLUS(me)
			    oidm.OpenAttachments("OPTIVA_REPORT_OUTPUT", _objectsymbol, _Objectkey, reportfilename)
			    
			    'shellapi("https://inforos107t.menz-gasser.it:9543/ca/api/items/OPTIVA_REPORT_OUTPUT-30-1-LATEST/resource/stream?%24language=en-US")
			    
            Else
                
                message = SO_gf.getValueJSON("message", responseString)
                If String.IsNullOrWhiteSpace(message) Then
                    message = SO_gf.getValueJSON("error", responseString)
                End If
                If String.IsNullOrWhiteSpace(message) Then
                    message = "Unknown failure. Raw response: " & responseString
                End If
                Throw New Exception("Failed to call API Error: [" & message & "]")
            End If
            
  
        End Using
        
    Catch ex as exception
    	messagelist("Error in " + _actioncode + ": " + ex.message)
    	PublishDebuggingInfo("Error: " + ex.message)
    	Return 9111
    end try
    
    PublishDebuggingInfo("Wf Generating IDM END")
    
    Return 111

End Function

End Class