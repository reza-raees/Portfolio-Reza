
' Purpose:  Print Formula's Batch Report  

' AUTHOR:   Reza
		
' NOTE: 	This file includes only the custom-developed logic for enhancing the Batch Formulas report.
' 			It does not contain full Interface Configuration or system templates generated.

	'========================================================================================
	' BEGIN: Adding data to the XML template for Bach Report 
	'========================================================================================
	'Get Root node
	Dim xnRoot As XMLNode
	xnRoot = xdTemplate.SelectSingleNode("/fsxml/report/object/FSFORMULA")
	
	'Date of the report
	Dim sDate as Date = DateTime.Now.Date
	
	Dim sDateNode as XMLNode
	sDateNode = xdTemplate.CreateElement("DATE")
	sDateNode.InnerText = sDate
	xnRoot.AppendChild(sDateNode)
	
	'creating nodes of the ingredients
	Dim sformCode as object = objproperty("FORMULA_CODE")
	'Dim sERPnumber as object = objproperty("ERP_FORMULA_ID")
	Dim ingritemCodes as object = objproperty("ITEMCODE.INGR", "", "", "*", 1)
	Dim ingrDescriotion as object = objproperty("DESCRIPTION.INGR", "", "", "*", 1)
	dim ingrActiveQua as object =objproperty("ACTIVEQUANTITY.INGR", "", "", "*", 1)
	Dim ingrQua as object = objproperty("ATTRIBUTE21.INGR.A", "", "", "*", 1)
	Dim ingruom as object = objproperty("UOMCODE.INGR", "", "", "*", 1)
	Dim ingrformula as object = objproperty("FORMULACODE.INGR","","","*",1)
	Dim mixnote as String = objproperty("DOCTEXT.DOC.A", "", "", "LAB_INSTRUCTIONS" ,1 )
	Dim totInput as object = objproperty("C_TOTAL_INPUT","","")
	Dim formYield as object = WipParamGet("BATCH_SIZE")
	Dim ingredInsruction as object = objproperty("INSTRUCTION.INGR", "", "", "*", 1)
	'Dim relationCalc as Object = formYield/totInput

	Dim itms as String = ""
	Dim descs as String = ""
	Dim actives as String = ""
	Dim qtys as String = ""
	Dim uoms as string = ""
	Dim forms as string = ""
	Dim totActqty as object = 0
	Dim totqty as object = 0
	Dim totalactinput as object = 0

	Dim ITEMNODE As XmlNode
	Dim DESCNODE As XmlNode
	Dim ACTNODE As XmlNode
	Dim QTYNODE As XmlNode
	Dim UOMNODE As XmlNode
	Dim FRMNODE As XmlNode 
	Dim INSTNOTENODE As XmlNode
	Dim EMPTYNOTE as XmlNode
	
	For x  as integer = 0 to ingritemCodes.length -1
		totalactinput = totalactinput + ingrActiveQua(x)
	Next
	
	Dim relationCalc as Object = formYield/totalactinput
	
	For x  as integer = 0 to ingritemCodes.length -1
		itms = ingritemCodes(x)
		descs = ingrDescriotion(x)
		actives = ingrActiveQua(x)
		qtys = ingrQua(x)
		uoms = ingruom(x)
		forms = ingrformula(x)

		'Create and append ROW node
		Dim rowNode As XmlNode = xdTemplate.CreateElement("ROW")
		xnRoot.AppendChild(rowNode)
		 
		'Dim instItemNote as string = objproperty("DOCTEXT.DOC.A", "ITEM", ingritemCodes(x) , "INSTRUC" ,1 )
		if IsBlank(ingredInsruction(x)) = 0 then
			EMPTYNOTE = xdTemplate.CreateElement("EMPTYNOTE")
			rowNode.AppendChild(EMPTYNOTE)
			
			INSTNOTENODE = xdTemplate.CreateElement("INSTNOTENODE")
			EMPTYNOTE.AppendChild(INSTNOTENODE)
			INSTNOTENODE.InnerText = ingredInsruction(x)
		end if 
	 
		ITEMNODE = xdTemplate.CreateElement("ITEMNODE")
		rowNode.AppendChild(ITEMNODE)
		ITEMNODE.InnerText = ingritemCodes(x)
		
		DESCNODE = xdTemplate.CreateElement("DESCNODE")
		rowNode.AppendChild(DESCNODE)
		DESCNODE.InnerText = ingrDescriotion(x)

		ACTNODE = xdTemplate.CreateElement("ACTNODE")
		rowNode.AppendChild(ACTNODE)
		ACTNODE.InnerText = ingrActiveQua(x) * relationCalc
		
		QTYNODE = xdTemplate.CreateElement("QTYNODE")
		rowNode.AppendChild(QTYNODE)
		QTYNODE.InnerText = ingrQua(x)
		
		UOMNODE = xdTemplate.CreateElement("UOMNODE")
		rowNode.AppendChild(UOMNODE)
		UOMNODE.InnerText = ingruom(x)
		
		FRMNODE = xdTemplate.CreateElement("FRMNODE")
		rowNode.AppendChild(FRMNODE)
		FRMNODE.InnerText = ingrformula(x)

		totActqty = totActqty + (ingrActiveQua(x) * relationCalc)
		totqty = totqty + ingrQua(x)
		
		'messagelist("item: " & ingritemCodes(x) &" and desc: " & ingrDescriotion(x) & " and act: " & ingrActiveQua(x) & " and qty:  " & ingrQua(x) & " and UOM: " & ingruom(x) )
		
	Next
	
	'total activeQua node
	Dim TOTACT As XmlNode
	TOTACT = xdTemplate.CreateElement("TOTACT")
	xnRoot.AppendChild(TOTACT)
	TOTACT.InnerText = totActqty
	
	'related Finished good node
	Dim FSHGNODE As XmlNode
	FSHGNODE = xdTemplate.CreateElement("FINISHEDGOOD")
	xnRoot.AppendChild(FSHGNODE)
	FSHGNODE.InnerText = WIPParamGet("MDMCODE")
	
	'total quantity 
	Dim TOTQUTY As XmlNode
	TOTQUTY = xdTemplate.CreateElement("TOTQUTY")
	xnRoot.AppendChild(TOTQUTY)
	TOTQUTY.InnerText = totqty
	
	'mixture note
	Dim MIXTURENOTE as XmlNode
	MIXTURENOTE = xdTemplate.CreateElement("MIXTURENOTE")
	xnRoot.AppendChild(MIXTURENOTE)
	MIXTURENOTE.InnerText = mixnote
		
	'========================================================================================
	' BEGIN: Adding data to the XML template for Bach Report
	'========================================================================================