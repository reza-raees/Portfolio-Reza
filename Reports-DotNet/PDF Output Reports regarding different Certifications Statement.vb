' Purpose:  Print Formula's specific parameters 

' AUTHOR:   Reza		June 2024 

' NOTE: 	This file includes only the custom-developed logic for creating different Certification Statements reports
' 			It does not contain full Interface Configuration or system templates generated.


Function wf_start() As Long
    
    Try
        
        Dim sXML as string
    	 sXML = ObjectXML("", "", "HEADER")
    
        'Library function handles post-ObjectXML updates and appends to xdTemplate
    	rptAdj.xmlAlterAppend(sXML, xdTemplate)
    	
    	'========================================================================================
        ' BEGIN - Add and update data into the reports' XML templates
        '========================================================================================
        'Get Root node
    	Dim xnRoot As XMLNode
        xnRoot = xdTemplate.SelectSingleNode("/fsxml/report/object/FSFORMULA")
        
        'Date of the report
        Dim sDate as Date = DateTime.Now.Date.ToString("MM/dd/yyyy")
        
        Dim sDateNode as XMLNode
        sDateNode = xdTemplate.CreateElement("SDATE")
        sDateNode.InnerText = sDate
        xnRoot.AppendChild(sDateNode)
    	
        'Bioengineered Statement
        Dim BIONODEYES as XMLNode
    	BIONODEYES = xdTemplate.CreateElement("BIONODEYES")
    	xnRoot.AppendChild(BIONODEYES)
    	
    	Dim BIONODENO as XMLNode
    	BIONODENO = xdTemplate.CreateElement("BIONODENO")
    	xnRoot.AppendChild(BIONODENO)
    	
    	Dim BEDERIVED as XMLNode
    	BEDERIVED = xdTemplate.CreateElement("BEDERIVED")
    	xnRoot.AppendChild(BEDERIVED)
    	
    	Dim BIONOTE as XMLNode
    	BIONOTE = xdTemplate.CreateElement("BIONOTE")
    	xnRoot.AppendChild(BIONOTE)
    	
    	Dim NOTEWRITE as XMLNode
    	NOTEWRITE = xdTemplate.CreateElement("NOTEWRITE")
        xnRoot.AppendChild(NOTEWRITE)
        
        Dim bioNotes as string = ObjProperty("ATTRIBUTE1.TPALL", "" , "" , "BE_STATUS" , "PARAM_CODE") 
        Dim biostatus as Object = ObjProperty("VALUE.TPALL", "" , "" , "BE_STATUS" , "PARAM_CODE")
        Dim bioderived as Object = ObjProperty("VALUE.TPALL", "" , "" , "BE_DERIVED" , "PARAM_CODE")
        
        if isblank(biostatus) = 0 and biostatus = "1" then
    	    BIONODEYES.InnerText = "X"
    	    BIONOTE.InnerText = bioNotes
    	    NOTEWRITE.InnerText = "BE Ingredients:"
	    elseif (biostatus = "0" and bioderived = "1") then
	        BIONODEYES.InnerText = ""
    	    BIONOTE.InnerText = ""
    	    NOTEWRITE.InnerText = ""
    	    BEDERIVED.InnerText = "X"
    	    BIONODENO.InnerText = ""
	    elseif isblank(biostatus) = 0 and biostatus = "0"
	        BIONODEYES.InnerText = ""
    	    BIONOTE.InnerText = ""
    	    NOTEWRITE.InnerText = ""
    	    BEDERIVED.InnerText = ""
    	    BIONODENO.InnerText = "X"
    	end if 
        
        'SUB_PROPP_65 Statement
    	Dim SUBS65NODEONE as XMLNode
    	SUBS65NODEONE = xdTemplate.CreateElement("SUBS65NODEONE")
    	xnRoot.AppendChild(SUBS65NODEONE)
        
    	Dim SUBS65NODEZERO as XMLNode
    	SUBS65NODEZERO = xdTemplate.CreateElement("SUBS65NODEZERO")
        xnRoot.AppendChild(SUBS65NODEZERO)
        
    	Dim NOTESUB65 as XMLNode
    	NOTESUB65 = xdTemplate.CreateElement("NOTESUB65")
        xnRoot.AppendChild(NOTESUB65)
        
    	if not WipParamGet("SUB65") is Nothing AndAlso WipParamGet("SUB65") = 1 then
    	    
    	    SUBS65NODEONE.InnerText = "X"
    	    SUBS65NODEZERO.InnerText = " "
    	    NOTESUB65.InnerText = WipParamGet("NOTESUB65")
	    else  
	        NOTESUB65.InnerText = ""
	        SUBS65NodEONE.InnerText = " "
	        SUBS65NODEZERO.InnerText = "X"
    	end if 
        
        'HALAL Statement
        Dim pHalal as Object = ObjProperty("VALUE.TPALL", "" , "" , "51_HALAL_SUITABLE" , "PARAM_CODE")
        Dim HALALNODEONE as XMLNode
    	HALALNODEONE = xdTemplate.CreateElement("HALALNODEONE")
    	xnRoot.AppendChild(HALALNODEONE)
        
    	Dim HALALNODEZERO as XMLNode
    	HALALNODEZERO = xdTemplate.CreateElement("HALALNODEZERO")
        xnRoot.AppendChild(HALALNODEZERO)
        
    	if pHalal = "1" then 
    	    HALALNODEONE.InnerText = "X"
    	    HALALNODEZERO.InnerText = " "
	    else  
	        HALALNODEONE.InnerText = " "
	        HALALNODEZERO.InnerText = "X"
    	end if 
        
        'VEGETARIAN Statement
        Dim pVagan as Object = ObjProperty("VALUE.TPALL", "" , "" , "VEGETARIAN_ROLLUP" , "PARAM_CODE")
        Dim OVOLACTOVEGNODE as XMLNode
    	OVOLACTOVEGNODE = xdTemplate.CreateElement("OVOLACTOVEG")
    	xnRoot.AppendChild(OVOLACTOVEGNODE)
    	
    	Dim LACTOVEGNODE as XMLNode
    	LACTOVEGNODE = xdTemplate.CreateElement("LACTOVEG")
    	xnRoot.AppendChild(LACTOVEGNODE)
    	
    	Dim VEGANNODE as XMLNode
    	VEGANNODE = xdTemplate.CreateElement("VEGAN")
    	xnRoot.AppendChild(VEGANNODE) 
        
        Dim NOTSUITABLENODE as XMLNode
    	NOTSUITABLENODE= xdTemplate.CreateElement("NOTSUITABLE")
    	xnRoot.AppendChild(NOTSUITABLENODE) 
        
        select case pVagan
            case "OVO_LACTO_VEG"
                OVOLACTOVEGNODE.InnerText = "X"
            
            case "LACTO_VEG"
                LACTOVEGNODE.InnerText = "X"
            
            case "VEGAN"
                VEGANNODE.InnerText = "X"
        
            case "UNK"
                NOTSUITABLENODE.InnerText = "X"
                
            Case "NON"
                NOTSUITABLENODE.InnerText = "X"
        end select
        
        'Genetically Engineered Statement
        Dim KNOWNODE as XMLNode
    	KNOWNODE = xdTemplate.CreateElement("KNOWNODE")
    	xnRoot.AppendChild(KNOWNODE)
    	
    	Dim NONGENNODE as XMLNode
    	NONGENNODE = xdTemplate.CreateElement("NONGENNODE")
    	xnRoot.AppendChild(NONGENNODE)
    	
    	Dim NONDERIVENODE as XMLNode
    	NONDERIVENODE = xdTemplate.CreateElement("NONDERIVENODE")
    	xnRoot.AppendChild(NONDERIVENODE)
    	
    	Dim GENENGNODE as XMLNode
    	GENENGNODE = xdTemplate.CreateElement("GENENGNODE")
    	xnRoot.AppendChild(GENENGNODE)
    	
        Dim gmrollup as Object = ObjProperty("VALUE.TPALL", "" , "" , "GMO_ROLLUP" , "PARAM_CODE")
        Dim gmderived as Object = ObjProperty("VALUE.TPALL", "" , "" , "55_GM_DERIV" , "PARAM_CODE")
        Dim genValue as Object = WipParamGet("GEN_VALUE")
        'messagelist("gmrollup: ",gmrollup)
        'messagelist("gmderived: ",gmderived)
        
        if genValue = "1" then
    	    KNOWNODE.InnerText = "X"
        end if
    	    
	    if gmrollup = "1" then 
    	    NONGENNODE.InnerText = "X"
	    end if
	        
        if gmderived = "0" then 
    	    NONDERIVENODE.InnerText = "X"
        end if
	            
	    if (gmrollup = "-1" and gmderived = "4") then
    	    GENENGNODE.InnerText = "X"
    	end if 
            
        '========================================================================================
        ' End - Add and update data into the reports' XML templates
        '========================================================================================

    
    Catch ex as exception
    	messagelist("Report Generation Error:  ", ex.message)
    End try

    return 111

End Function

