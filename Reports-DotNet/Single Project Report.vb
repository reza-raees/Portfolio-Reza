' Workflow Action: SO_RPT_SINGLE_PROJECT_DATA 
' May 2025 

' Purpose:  Print Project's Data

' AUTHOR    :   Reza
' REVISIONS :   WHO         DATE CHANGE     REASON

'               ---------   ------------    ------------------------------------------------------
Option Strict Off
imports System
imports System.Diagnostics
imports System.XML
Imports Microsoft.VisualBasic
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data
Imports System.IO
Imports System.Text
Imports System.Data.DataRelation
Imports DocICP = Infor.DocumentManagement.ICP
Imports System.Text.RegularExpressions
Imports System.Data.Common
Imports System.Globalization
Imports System.Net
Imports System.Net.Http
Imports System.Net.Http.Headers
Imports System.Net.Mime
Imports System.Web

Class ActionScript
 Inherits FcProcFuncSetEventWF

Function wf_start() As Long
    
    Try
    
        '================
        'Libraries in use:
    	dim rptAdj as SO_REPORTADJ = New SO_REPORTADJ(Me)
    	dim ck as CLEANKEY = new CLEANKEY(Me)
    	DIM idmReportHelper as IDMREPORTHELPER = New IDMREPORTHELPER(Me)
    	
        '================
        'Define doc code here - required if using IDM Reports; Optional for WebReports
    	dim doccode as string = "OPTIVA_REPORT_OUTPUT"
    	
        'Define Report template name according to the actionset name for IDM Reports (Must be same as template file Title in IDM):
        
        dim reportname as string = "Sample IDM Template"
	        
    	dim extension as string = ".docx"	' change extension to docx to generate as Word doc
    	dim filename as string = reportname & "_" & _objectkey  & "_" & DateTime.Now.ToString("yyyyMMdd_HH_mm_ss") & extension
        '================
        'Part 1 - CHECK VALUES!  initial declarations and file paths
        '==============
        'Specify the name of the Optiva XML data file.
    	dim job as string = "SO_FORMULA_LABEL"  
    
        'Define the report output type (use PDF, MSWORD, OR EXCEL):
    	dim reporttype as string = "MSWORD"         
    
        'Do you want the output saved to a specific file path and name? Enter YES or NO in capital letters
    	dim saveReport as string = "NO"           
    	'If saveReport is YES, will the report be attached to a Function Code on the primary object?
    	dim attachReport as string = "NO"
    
        'Create unique XML and output file names
    	Dim uid as string = ck.update(_objectkey)            ' remove \, *, and other special characters
    	dim jobfileName as string = job + "_" + uid + "_" + cstr(_wipid) + ".xml"              'modify as needed to create a unique filename for the XML
    	dim outputfile as string = job + "_" + uid + "_" + cstr(_wipid) + rptAdj.fileext(reporttype)
    
        '================
        'Part 2 - CHECK VALUES!  define schema list
        '================
        'Create a template XML file to contain the report data. Include the schema names to use for the report.
        'List all schema files needed for the report within the curly brackets, in double quotes, seperated by commas e.g. {"Schema1.xsd", "Schema2.xsd", ...}
    
    	dim schemaList() as string = {"PROJECT_IA.xsd"}
    	dim sTemplate as string = rptAdj.BuildTemplate(schemaList)	'Runs Library script
    
        'Load the template string to begin the XML Document
    	dim xdTemplate as New XmlDocument
    	xdTemplate.LoadXml(sTemplate)
  
        '================
        'Part 3 - CHECK VALUES!  define the XML details for the object(s) in question... use only what is needed.
        '================
        Dim sXML as string
        'Create an XML file for the object. Specify details to include.  
    	'sXML = ObjectXML("", "", "HEADER;ST;DOC;ATTACH;PER;STATUS;VIEWS;REF;CONTEXT;TPALL;CUSTOM;MATRIX;HIST;HISTSUMMARY")
    	 sXML = ObjectXML("", "", "HEADER;DOC;CONTEXT;TPALL;MATRIX")
    
        'Library function handles post-ObjectXML updates and appends to xdTemplate
    	rptAdj.xmlAlterAppend(sXML, xdTemplate)
    	'=====================================
        ' Adding Custom data to the XML - BEGIN
        '=====================================
        'Get Root node
    	Dim xnRoot As XMLNode
        xnRoot = xdTemplate.SelectSingleNode("/fsxml/report/object/FSPROJECT")
        
        
       'Date of the report
        Dim ssDate As Date = DateTime.Now.Date
        Dim sDate As String = ssDate.ToString("MM/dd/yyyy")
       
        Dim xmlNodex As XMLNode
        Dim sDateNode As XmlNode
        sDateNode = xdTemplate.CreateElement("SDATE")
        sDateNode.InnerText = sDate
        xnRoot.AppendChild(sDateNode)
        
        'getEnumLabel("C_YES_NO_UNKNOWN", row("PVALUE"))
       ' 'VOLUMI STIMATI and SCHEMA ARTICOLI nodes
       ' Dim volStimatifield2  as Object = ObjProperty("FIELD2.MATRIX.V\DMPROJECT4","","","*")
       ' Dim schemaArticoli2  as Object = ObjProperty("FIELD2.MATRIX.V\DMPROJECT3","","","*")
       ' 
       ' xmlNodex = xdTemplate.CreateElement("VOLUMI_STIMATI_TABLE")
       ' xnRoot.AppendChild(xmlNodex)
       ' xmlNodex.InnerText =  ""
       ' for i as integer = 0 to volStimatifield2.length -1
       '     If volStimatifield2(i) <> "" Then
       '         xmlNodex.InnerText =  "1"
       '        'messagelist(volStimatifield2(i))
       '     end if 
       ' Next
       ' 'messagelist(xmlNodex.InnerText)
       ' 
       ' xmlNodex = xdTemplate.CreateElement("SCHEMA_ARTICOLI_TABLE")
       ' xnRoot.AppendChild(xmlNodex)
       'If schemaArticoli2.Length > 0 Then
       '    xmlNodex.InnerText =  "1"
       'else 
       '    xmlNodex.InnerText =  ""
       'end if
       ' messagelist(schemaArticoli2.Length.ToString())
        
        
    	Dim ds as DataSet = New DataSet()
    	Dim reader As StringReader = New StringReader(xdTemplate.outerxml)
    	ds.Readxml(reader)
    
    	Dim row As DataRow
    	Dim xmlData as String
    	
    	'Adding Enum Label of the fields
    	Dim formatoPrimarioLabel As String = ""
        Dim formatoRiconfezionamentoLabel As String = ""
        Dim formatoSecondarioLabel As String = ""
            
    	'If ds.Tables.Contains("FSPROJEC") Then
	    For Each row In ds.Tables("FSPROJECT").Rows
            If row.Table.Columns.Contains("PROD_INTRO_DATE") AndAlso Not IsDBNull(row("PROD_INTRO_DATE")) Then
                Dim rawDate As DateTime = Convert.ToDateTime(row("PROD_INTRO_DATE"))
                Dim formattedDate As String = rawDate.ToString("dd/MM/yyyy", New CultureInfo("it-IT"))
                'messagelist(formattedDate)
                
                'ds.Tables("FSPROJECT").Columns.Add("PROD_INTRO_DATE_UPDATE", GetType(String))
                row("PROD_INTRO_DATE") = formattedDate
            End If  
            
            If row.Table.Columns.Contains("PROJECTMGR") AndAlso Not IsDBNull(row("PROJECTMGR")) Then
                row("PROJECTMGR") = getEnumLabel("PROJECTMGR",row("PROJECTMGR"))
               ' messagelist(row("PROJECTMGR"))
            End If  
              
	    Next
	    'End if 
    	
    	'Extracting Project Params
    	If ds.Tables.Contains("FSPROJECTTPALL") Then
            ds.Tables("FSPROJECTTPALL").Columns.Add("VALUE_EXT", GetType(String))
        
            Dim pValueLabel As String
            Dim pValueFormatted As String
            Dim pValue As String
        
            For Each row In ds.Tables("FSPROJECTTPALL").Rows
                If row.Table.Columns.Contains("PVALUE_LABEL") Then
                    pValueLabel = If(IsDBNull(row("PVALUE_LABEL")), String.Empty, row("PVALUE_LABEL"))
                End If
                If row.Table.Columns.Contains("PVALUE_FORMATTED") Then
                    pValueFormatted = If(IsDBNull(row("PVALUE_FORMATTED")), String.Empty, row("PVALUE_FORMATTED"))
                End If
                If row.Table.Columns.Contains("PVALUE") Then
                    pValue = If(IsDBNull(row("PVALUE")), String.Empty, row("PVALUE"))
                End If
                row("VALUE_EXT") = GetPValue(pValueLabel, pValueFormatted, pValue)
            Next
    	End If
    	
       'Extracting EXTENSION TABELS retrivieng  
       Dim rowsToRemove As New List(Of DataRow)

        For Each rowi As DataRow In ds.Tables("FSPROJECTMATRIX_4").Rows
            If (Not rowi.Table.Columns.Contains("FIELD2") OrElse IsDBNull(rowi("FIELD2"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD3") OrElse IsDBNull(rowi("FIELD3"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD4") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD5") OrElse IsDBNull(rowi("FIELD5"))) Then
                
                'messagelist("Removed: ", rowi("FIELD1"))
                rowsToRemove.Add(rowi) 
            End If
        Next
        
        For Each row In rowsToRemove
            ds.Tables("FSPROJECTMATRIX_4").Rows.Remove(row)
        Next
        
        'Extracting EXTENSION TABELSSCHEMA ARTICOLI
        Dim rwsToRemove As New List(Of DataRow)
        For Each rowi As DataRow In ds.Tables("FSPROJECTMATRIX_3").Rows
            If (Not rowi.Table.Columns.Contains("FIELD2") OrElse IsDBNull(rowi("FIELD2"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD3") OrElse IsDBNull(rowi("FIELD3"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD4") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD5") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD6") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD7") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD8") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD9") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD10") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD11") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD12") OrElse IsDBNull(rowi("FIELD4"))) AndAlso
               (Not rowi.Table.Columns.Contains("FIELD13") OrElse IsDBNull(rowi("FIELD4"))) Then
                
                'messagelist("Removed: ", rowi("FIELD1"))
                rwsToRemove.Add(rowi) 
            End If
        Next
        
        ' Now remove rows outside the loop
        For Each row In rwsToRemove
            ds.Tables("FSPROJECTMATRIX_3").Rows.Remove(row)
        Next
        
       'Extracting Project Notes
        If ds.Tables.Contains("FSPROJECTDOC") Then
            Dim dt As DataTable = ds.Tables("FSPROJECTDOC")
            For i As Integer = dt.Rows.Count - 1 To 0 Step -1
                Dim rowd As DataRow = dt.Rows(i)
                Dim funcCode As String = rowd("FUNCTION_CODE").ToString().Trim()
                Dim textData As String = If(IsDBNull(rowd("TEXT_DATA")), "", rowd("TEXT_DATA").ToString().Trim())
        
                If funcCode = "HISTORY" OrElse (funcCode <> "HISTORY" AndAlso String.IsNullOrWhiteSpace(textData)) Then
                    'messagelist(funcCode)
                    dt.Rows.RemoveAt(i)
                End If
            Next
        End If

            
        '=====================================
        ' Adding Custom data to the XML - END
        '=====================================

    	'Create string of updated XML
    	xmlData = ds.GetXml()
    	
    	
    	'=====================================
        'Generate and Launch the Report by calling another wf to call the API to launch the report 
        '=====================================
        
    	'***** GENERATE AND LAUNCH REPORT *****
		'for testing or build template, save XML:
		' dim xmlsave as string = SaveReportData(xmlData)
		' messagelist("XML Saved to ", xmlsave)
		' messagelist("XML:")
		'messagelist(xmlData)
		'messagelist("in the report script xml report: ", reportxmldatabase64)
		'messagelist("in the filename xml report: ",filename)
		
		filename = filename.Replace("\", "_")
		'messagelist(filename)
        StartWorkflow("SO_IDM_REPORT_CALLING_API","","", xmlData , filename)
    
    Catch ex as exception
    	messagelist("Report Generation Error:  ", ex.message)
    End try

    return 111

End Function

'***************************************************************************************
Function GetPValue(pValueLabel As String, pvalueFormatted As String, pvalue As String) As String
	Dim pVal As String = String.Empty
	If Not String.IsNullOrEmpty(pValueLabel) And Len(Trim(pValueLabel)) > 0 Then
		pVal = pValueLabel
	ElseIf Not String.IsNullOrEmpty(pvalueFormatted) Then
		pVal = pvalueFormatted
	ElseIf Not String.IsNullOrEmpty(pvalue) And Len(Trim(pvalue)) > 0 Then
		pVal = pvalue
	Else pVal = ""
	End If
	Return pVal
End Function

Function getEnumLabel(byval enumList as string, byval enumValue as string) as string
    Dim dsList As DataSet = ObjectDataSet("ENUMLIST", enumList, "CF")
	Dim EnumTableName As String = DataSetTableName("ENUMLIST", enumList, "CF")
	'MessageList("EnumTableName: " & EnumTableName)
	Dim drVal() As Data.DataRow = dsList.Tables(EnumTableName).Select("ENUM_VALUE = '"+ enumValue +"'")
	Dim enumlabel As String = " || "
	If drVal IsNot Nothing And drVal.Length > 0 Then
        enumlabel = drVal(0)("ENUM_LABEL2")
	End if

    return enumlabel
End Function


End Class
