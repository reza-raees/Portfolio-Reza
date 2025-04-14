' NOTE:  These are the customised Developments over different Reports 
' 			It does not contain full Interface Configuration or system templates generated.
 

'===========================================================================
' begin: Customisation for specification report (Adding Footer dates, Sites, for both formulas and its alternatives)
'===========================================================================
Dim xnRoot As XMLNode
xnRoot = xdTemplate.SelectSingleNode("/fsxml/report/object/FSFORMULA")

'Date of the report
Dim ssDate as Date = DateTime.Now.Date
Dim revisionDate as string = ssDate.ToString("dd/MM/yyyy")
Dim revisiondateNode as XMLNode
revisiondateNode = xdTemplate.CreateElement("REVISIONDATE")
revisiondateNode.InnerText = revisionDate
xnRoot.AppendChild(revisiondateNode)
'messagelist("revis: ", revisionDate)
'messagelist("SSDATE:",ssDate)

Dim sDate as string = ssDate.ToString("dd/MM/yyyy")

'Issue Date of the report
Dim wfDateTable As DataTable = TableLookupEx("SO_SPEC_ISSUE_DATE_IS_EU", "TableLookup", _ACTIONCODE, _OBJECTKEY)
'messagelist(_OBJECTKEY)

If wfDateTable.Rows.Count > 0 Then
    ' Access only the first row
    Dim firstRow As DataRow = wfDateTable.Rows(0)
    Dim parsedDate As DateTime
    Dim formattedDate As String = String.Empty

    ' Attempt to parse the date from the column OLDEST_DATE
    If DateTime.TryParse(firstRow("COMPLETION_DATE").ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, parsedDate) Then
        formattedDate = parsedDate.ToString("dd/MM/yyyy")
        'messagelist("Formatted Date: " & formattedDate)
        sDate = formattedDate
    End If

End If

Dim sDateNode as XMLNode
sDateNode = xdTemplate.CreateElement("SDATE")
xnRoot.AppendChild(sDateNode)
sDateNode.InnerText = sDate
'messagelist("SDATE:", sDate)

'adding notes to the product details
Dim legalNote as String = ObjProperty("DOCTEXT.DOC.A", "", "", "LEGAL NAME/DENOMINATION" ,1 )
if IsBlank(legalNote) =  0 then
    Dim legalNode As XMLNode
    Dim valueLegalNote as XmlNode
    legalNode = xdTemplate.CreateElement("LEGALNOTE")
    xnRoot.AppendChild(legalNode)
    
    valueLegalNote = xdTemplate.CreateElement("NOTELEGAL")
    xnRoot.AppendChild(valueLegalNote)
    valueLegalNote.InnerText = legalNote
end if

Dim microbionote as string = objproperty("DOCTEXT.DOC.A", "", "", "MICROBIOLOGICAL_STANDARDS" ,1)
if IsBlank(microbionote) = 0 then
    Dim MicroNode as XmlNode
    MicroNode = xdTemplate.CreateElement("MICRONOTE")
    xnRoot.AppendChild(MicroNode)
end if 

Dim typicalvaluesNote as string = objproperty("DOCTEXT.DOC.A", "", "", "TYPICAL_VALUES" ,1)
if IsBlank(typicalvaluesNote) = 0 then
    Dim typicalNode as XmlNode
    typicalNode = xdTemplate.CreateElement("TYPICALNOTE")
    xnRoot.AppendChild(typicalNode)
end if 

'adding the Country of Origin and languagenode regarding the site 
Dim languageNode as XmlNode
languageNode = xdTemplate.CreateElement("LANGUAGE_REPORT")
xnRoot.AppendChild(languageNode)

Dim siteNodes as XmlNode

Dim sAttribcodeContext As Object = ObjProperty("ATTRIBCODE.CONTEXT", "", "", "*", 1)
Dim sWorkcenter As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)

Dim countryOriginList As New List(Of String)
For x As Integer = 0 To sWorkcenter.Length - 1
    'messagelist(sAttribcodeContext(x))
    if (sAttribcodeContext(x) = "MFGLOC" or sAttribcodeContext(x) = "SELLOC") then
        'messagelist("workcenter: ", sWorkcenter(x))
        Dim y As String = If(x = 0, "A", If(x = 1, "B", If(x = 2, "C", If(x = 3, "D", If(x = 4, "E", "")))))
        Dim sWarehouse As String = ObjProperty("PARENT_CODE", "LOCATION", sWorkcenter(x))
        siteNodes = xdTemplate.CreateElement("COUNTRY" + y)
        xnRoot.AppendChild(siteNodes)
        Select Case True
            Case sWarehouse.StartsWith("B")
                If Not countryOriginList.Contains("UNITED KINGDOM") Then
                    If countryOriginList.Count = 0 Then
                        siteNodes.InnerText = "UNITED KINGDOM"
                    Else
                        siteNodes.InnerText = " - UNITED KINGDOM"
                    End If
                    countryOriginList.Add("UNITED KINGDOM") 
                End If
                
            Case sWarehouse.StartsWith("G")
                If Not countryOriginList.Contains("GERMANY") Then
                    If countryOriginList.Count = 0 Then
                        siteNodes.InnerText = "GERMANY"
                    Else
                        siteNodes.InnerText = " - GERMANY"
                    End If
                    countryOriginList.Add("GERMANY")
                End If
                
            Case sWarehouse.StartsWith("F")
                If Not countryOriginList.Contains("FRANCE") Then
                    If countryOriginList.Count = 0 Then
                        siteNodes.InnerText = "FRANCE"
                    Else
                        siteNodes.InnerText = " - FRANCE"
                    End If
                    countryOriginList.Add("FRANCE") 
                End If
                
            Case sWarehouse.StartsWith("N")
                If Not countryOriginList.Contains("NETHERLANDS") Then
                    If countryOriginList.Count = 0 Then
                        siteNodes.InnerText = "NETHERLANDS"
                    Else
                        siteNodes.InnerText = " - NETHERLANDS"
                    End If
                    countryOriginList.Add("NETHERLANDS") 
                End If
                
            Case sWarehouse.StartsWith("I")
                If Not countryOriginList.Contains("ITALY") Then
                    If countryOriginList.Count = 0 Then
                        siteNodes.InnerText = "ITALY"
                    Else
                        siteNodes.InnerText = " - ITALY"
                    End If
                    countryOriginList.Add("ITALY") 
                End If  
        end Select
        if x = 0 then
            If countryOriginList.Contains("ITALY") Then
                languageNode.InnerText = "IT"
            else 
                languageNode.InnerText = "EN"
            end if 
        end if 
    end if 
Next

'trying to do the same for alternative formual
Dim AltSiteNode as XmlNode
Dim formCode as string = ObjProperty("FORMULA_CODE")
Dim ForCode as Object = ObjProperty("KEYCODE", "", "").Split("-")(0)
Dim VersionFor as String = ObjProperty("VERSION", "", "")
Dim altFormulas as DataTable = TableLookupEx("SO_RPT_ALTER_FRM_IS_EU","TableLookup",ForCode,VersionFor)
Dim n as Integer = 0
for each row as datarow in altFormulas.Rows 
    n = n + 1
    Dim m As String = If(n = 1, "F", If(n = 2, "G", If(n = 3, "H", If(n = 4, "I", If(n = 5, "J", "")))))
    Dim altFormula as string = row("FORMULA_CODE")
    Dim objKy as string = row("FORMULA_CODE") +"\"+ row("VERSION")
    'messagelist(altFormula)
    if altFormula <> formCode then
        
        Dim altWorkcenter As Object = ObjProperty("ATTRIBVAL.CONTEXT", "FORMULA", objKy , "*", 1)
        Dim altAttribcodeContext As Object = ObjProperty("ATTRIBCODE.CONTEXT", "FORMULA", objKy, "*", 1)

        For x As Integer = 0 To altWorkcenter.Length - 1
            if (altAttribcodeContext(x) = "MFGLOC" or altAttribcodeContext(x) = "SELLOC") then
                Dim y As String = If(x = 0, "A", If(x = 1, "B", If(x = 2, "C", If(x = 3, "D", If(x = 4, "E", "")))))
                'messagelist(altWorkcenter(x))
                Dim altWarehouse As String = ObjProperty("PARENT_CODE", "LOCATION", altWorkcenter(x))
                AltSiteNode = xdTemplate.CreateElement("ALTCOUNTRY"+y+m)
                xnRoot.AppendChild(AltSiteNode)
                        
                Select Case True
                Case altWarehouse.StartsWith("B")
                    If Not countryOriginList.Contains("UNITED KINGDOM") Then
                        AltSiteNode.InnerText = "- UNITED KINGDOM"
                        countryOriginList.Add("UNITED KINGDOM") 
                    End If
                    
                Case altWarehouse.StartsWith("G")
                    If Not countryOriginList.Contains("GERMANY") Then
                        AltSiteNode.InnerText = "- GERMANY"
                        countryOriginList.Add("GERMANY") 
                    End If
        
                Case altWarehouse.StartsWith("F")
                    If Not countryOriginList.Contains("FRANCE") Then
                        AltSiteNode.InnerText = "- FRANCE"
                        countryOriginList.Add("FRANCE") 
                    End If
        
                Case altWarehouse.StartsWith("N")
                    If Not countryOriginList.Contains("NETHERLANDS") Then
                        AltSiteNode.InnerText = "- NETHERLANDS"
                        countryOriginList.Add("NETHERLANDS") 
                    End If
                    
                Case altWarehouse.StartsWith("I")
                    If Not countryOriginList.Contains("ITALY") Then
                        AltSiteNode.InnerText = "- ITALY"
                        countryOriginList.Add("ITALY")
                    End If
                End Select
            end If
        next
    end If
next

'times the report launched with the tranlation mode
Dim nxtkeyNumTran as object 
dim objCode as String = ObjProperty("FORMULA_CODE") + "\" + ObjProperty("VERSION")
Dim pubDateTable as DataTable = TableLookupEx("SO_RPT_PUBLISHED_TO_GMT_IS_EU","TableLookup", objCode , _ACTIONCODE)
if pubDateTable.Rows.Count = 0 then
    nxtkeyNumTran = 1
else
    nxtkeyNumTran = pubDateTable.Rows.Count + 1
end if 

Dim NextKeyNumberNode as XmlNode
NextKeyNumberNode = xdTemplate.CreateElement("NEXT_KEY_NUMBER_TRANSLATED")
xnRoot.AppendChild(NextKeyNumberNode)
NextKeyNumberNode.InnerText = cstr(nxtkeyNumTran)

'times the report launched for the current object
Dim nxtkeyNumLaun as object 
Dim launchedDateTable as DataTable = TableLookupEx("SO_RPT_TDS_LAUNCHED_IS_EU","TableLookup", objCode , _ACTIONCODE)
if launchedDateTable.Rows.Count = 0 then
    nxtkeyNumLaun = 1
else
    nxtkeyNumLaun = launchedDateTable.Rows.Count + 1
end if 

Dim NextKeyLaunchedNode as XmlNode
NextKeyLaunchedNode = xdTemplate.CreateElement("NEXT_KEY_NUMBER_LAUNCHED")
xnRoot.AppendChild(NextKeyLaunchedNode)
NextKeyLaunchedNode.InnerText = cstr(nxtkeyNumLaun)


Dim tdsOption as long = WipParamGet("CUSTOMER_PUBLISH")

Dim PublishTDSOption as XmlNode
PublishTDSOption = xdTemplate.CreateElement("Publish_TDS_Option")
xnRoot.AppendChild(PublishTDSOption)
PublishTDSOption.InnerText = cstr(tdsOption)

'adding workflow name
Dim ReportNode as XmlNode
ReportNode = xdTemplate.CreateElement("REPORT_NAME")
xnRoot.AppendChild(ReportNode)
ReportNode.InnerText = _ACTIONSETCODE

'===========================================================================
'End: Customisation for specification report (Adding Footer dates, Sites, for both formulas and its alternatives)
'===========================================================================


'===========================================================================
'Begin: getting the botonical origin for all RM ingredients 
'===========================================================================

If (ds.Tables.Contains("FSFORMULALINEEXP")) Then 

	' Ensure column exists
	If Not ds.Tables("FSFORMULALINEEXP").Columns.Contains("C_ORIGINE_BOTANICA") Then
		ds.Tables("FSFORMULALINEEXP").Columns.Add("C_ORIGINE_BOTANICA", GetType(String))
		'messagelist("Added column C_ORIGINE_BOTANICA to FSFORMULALINEEXP.")
	End If

	' Process each row in FSFORMULALINEEXP
	For Each row In ds.Tables("FSFORMULALINEEXP").Rows
		If Not IsDBNull(row("CLASS")) AndAlso Not String.IsNullOrEmpty(row("CLASS").ToString()) AndAlso row("CLASS").ToString() = "RM" Then
			
			' Load item dataset
			Dim dsTest As DataSet = ObjectDataSet("ITEM", row("ITEM_CODE"), "HEADER")
			
			If dsTest.Tables.Contains("FSITEM") AndAlso dsTest.Tables.Contains("FSITEM_CUSTOMATTRIB") Then
				Dim itemAttRows As DataRow() = dsTest.Tables("FSITEM_CUSTOMATTRIB").Select("ITEM_CODE = '" & row("ITEM_CODE").ToString() & "' AND ATTRIB_CODE = 'C_ORIGINE_BOTANICA'")

				If itemAttRows.Length > 0 Then
					Dim cOrigineBotanica As String = ""

					' Aggregate C_ORIGINE_BOTANICA values
					For Each attribRow As DataRow In itemAttRows
						If Not IsDBNull(attribRow("ATTRIB_VAL")) AndAlso Not String.IsNullOrEmpty(attribRow("ATTRIB_VAL").ToString()) Then
							Dim label As String = getEnumLabel("C_COUNTRIES2", attribRow("ATTRIB_VAL").ToString())
							If String.IsNullOrEmpty(cOrigineBotanica) Then
								cOrigineBotanica = label
							Else
								cOrigineBotanica &= "; " & label
							End If
						Else
							messagelist("ATTRIB_VAL is missing for ITEM_CODE: " & row("ITEM_CODE").ToString())
						End If
					Next

					' Assign to column
					row("C_ORIGINE_BOTANICA") = cOrigineBotanica
				'Else
					'messagelist("No C_ORIGINE_BOTANICA found for ITEM_CODE: " & row("ITEM_CODE").ToString())
				End If
			'Else
				'messagelist("Missing FSITEM or FSITEM_CUSTOMATTRIB table for ITEM_CODE: " & row("ITEM_CODE").ToString())
			End If
		End If
	Next
End If

'===========================================================================
'' Purpose:  Generating XML file of a Formula and Share it via Interface 
End : getting the botonical origin for all RM ingredients 
'===========================================================================