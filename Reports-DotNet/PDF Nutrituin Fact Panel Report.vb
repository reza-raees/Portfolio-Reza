' Purpose:  Print Formula's Nutrituin Panel Facts

' AUTHOR    :   Reza     July 2024

' NOTE: 	This file includes only the custom-developed logic for creating the Nutrition Facts Panel report needed parameters and calculations.
' 			It does not contain full Interface Configuration or system templates generated.

Function wf_start() As Long
    
    Try
    
        Dim sXML as string
    	 sXML = ObjectXML("", "", "HEADER")
    
        'Library function handles post-ObjectXML updates and appends to xdTemplate
    	rptAdj.xmlAlterAppend(sXML, xdTemplate)
    	
    	'========================================================================================
        ' Begin - Add Samples data to the report XML template for NFP
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
    	
    	'retrieving the initial info of workflow
    	Dim houseValue as Double = WipParamGet("HOUSEHOLD")
		Dim HOUSEHOLD As XMLNode
        HOUSEHOLD = xdTemplate.CreateElement("HOUSEHOLD")
        xnRoot.AppendChild(HOUSEHOLD)
        HOUSEHOLD.InnerText = WipParamGet("HOUSEHOLD")
        
        Dim UNITSIZE As XMLNode
        UNITSIZE = xdTemplate.CreateElement("UNITSIZE")
        xnRoot.AppendChild(UNITSIZE)
        UNITSIZE.InnerText = WipParamGet("UNIT_SIZE")
        
        Dim NUMSERV As XMLNode
        NUMSERV = xdTemplate.CreateElement("NUMSERV")
        xnRoot.AppendChild(NUMSERV)
        NUMSERV.InnerText = WipParamGet("NUM_SERV")
        
        'Initialize the amount per serving and rounding the values and their xml nodes
        'calories
        Dim ServingCalories as object
        Dim CaloriesValue as Object = ObjProperty("VALUE.TPALL", "" , "", "ENERGY_KCAL_AMER", "PARAM_CODE")
        if IsBlank(CaloriesValue) = 0 AndAlso Not IsDBNull(CaloriesValue) then
            ServingCalories = ( CaloriesValue * houseValue) / 100
        end if
        
        Dim roundCaloris as object 
        if IsBlank(ServingCalories) = 0 AndAlso Not IsDBNull(ServingCalories) then
            roundCaloris = RoundValueGreaterOne(ServingCalories, 5 , 50 , 5 , 10)
            'messagelist("serv Value calories is: " & ServingCalories & " and rounded is: "&roundCaloris)
        else 
            roundCaloris = 0
        end if
        Dim SERVCALORIES As XmlNode
    	SERVCALORIES = xdTemplate.CreateElement("SERVCALORIES")
        xnRoot.AppendChild(SERVCALORIES)
        SERVCALORIES.InnerText = roundCaloris
        
        'fat
        Dim ServingFat as object
        Dim fatValue as object = ObjProperty("VALUE.TPALL", "","", "FAT" , "PARAM_CODE")
        if IsBlank(fatValue) = 0 AndAlso Not IsDBNull(fatValue) then 
            ServingFat = ( fatValue * houseValue) / 100
        end if
        
        Dim roundFat as object
        If Not IsBlank(ServingFat) AndAlso Not IsDBNull(ServingFat) AndAlso IsNumeric(ServingFat) Then
            roundFat = RoundValueHalfOne(CDbl(ServingFat), 0.5, 5)
        Else
            ServingFat = 0.0
            roundFat = 0.0
        End If
        'messagelist("serv Value fat is: " & ServingFat & " and rounded is: " & roundFat)
        Dim SERVFAT As XmlNode
    	SERVFAT = xdTemplate.CreateElement("SERVFAT")
        xnRoot.AppendChild(SERVFAT)
        SERVFAT.InnerText = roundFat.ToString()
        
        'fatsat
        Dim ServingFatSat as object
        Dim fatSatValue as object = ObjProperty("VALUE.TPALL", "" , "" , "FASAT_G" , "PARAM_CODE")
        if IsBlank(fatSatValue) = 0 AndAlso Not IsDBNull(fatSatValue) then 
            ServingFatSat = ( fatSatValue * houseValue) / 100
        end If
        
        Dim roundFatSat as object 
        If Not IsBlank(ServingFatSat) AndAlso Not IsDBNull(ServingFatSat) AndAlso IsNumeric(ServingFatSat) Then
            roundFatSat = RoundValueHalfOne(CDbl(ServingFatSat), 0.5, 5)
        Else
            ServingFatSat = 0.0
            roundFatSat = 0.0
        End If
        'messagelist("serv Value fatsat is: " & ServingFatSat & " and rounded is: "&roundFatSat)
        Dim SERVFATSAT As XmlNode
    	SERVFATSAT = xdTemplate.CreateElement("SERVFATSAT")
        xnRoot.AppendChild(SERVFATSAT)
        SERVFATSAT.InnerText = roundFatSat.ToString()
        
        'fattrn
        Dim ServingFatTrn as object
        Dim fatTrnValue as object = ObjProperty("VALUE.TPALL", "" , "", "FATRN_G" , "PARAM_CODE")
        if IsBlank(fatTrnValue) = 0 AndAlso Not IsDBNull(fatTrnValue) then 
            ServingFatTrn = ( fatTrnValue * houseValue) / 100
        end If
        
        Dim roundFatTrn as object  
        If Not IsBlank(ServingFatTrn) AndAlso Not IsDBNull(ServingFatTrn) AndAlso IsNumeric(ServingFatTrn) Then
            roundFatTrn = RoundValueHalfOne(CDbl(ServingFatTrn), 0.5, 5)
        Else
            ServingFatTrn = 0.0
            roundFatTrn = 0.0
        End If
        'messagelist("serv Value fatTrn is: " & ServingFatTrn & " and rounded is: "&roundFatTrn)
        Dim SERVFATTRN As XmlNode
    	SERVFATTRN = xdTemplate.CreateElement("SERVFATTRN")
        xnRoot.AppendChild(SERVFATTRN)
        SERVFATTRN.InnerText = roundFatTrn.ToString()
        
        'cholestrole
        Dim ServingChole as object
        Dim cholValue as object = ObjProperty("VALUE.TPALL", "" , "", "CHOLE_MG" , "PARAM_CODE")
        if IsBlank(cholValue) = 0 AndAlso Not IsDBNull(cholValue) then
            ServingChole = (cholValue * houseValue) / 100
        end if
        
        Dim roundChole as object
        If Not IsBlank(ServingChole) AndAlso Not IsDBNull(ServingChole) AndAlso IsNumeric(ServingChole) Then
            If ServingChole < 2 Then
                roundChole = 0
            ElseIf ServingChole < 5 Then
                roundChole = "<5"
            Else
                roundChole = Math.Round(ServingChole/ 5 , MidpointRounding.AwayFromZero) * 5
            End If
        Else
            ServingChole = 0.0
            roundChole = 0.0
        End If
        'messagelist("serv Value Chole is: " & ServingChole & " and rounded is: "&roundChole)
        Dim SERVCHOLE As XmlNode
    	SERVCHOLE = xdTemplate.CreateElement("SERVCHOLE")
        xnRoot.AppendChild(SERVCHOLE)
        SERVCHOLE.InnerText = roundChole.ToString()
        
        'Total Carb
        Dim ServingCarb as object
        Dim carbValue as object = ObjProperty("VALUE.TPALL", "" , "", "CARB_TOTAL_AMER" , "PARAM_CODE") 
        if IsBlank(carbValue) = 0 AndAlso Not IsDBNull(carbValue) then
            ServingCarb = (carbValue * houseValue) / 100 
        end if 
        
        Dim roundCarb as Object  
        If Not IsBlank(ServingCarb) AndAlso Not IsDBNull(ServingCarb) AndAlso IsNumeric(ServingCarb) Then
            roundCarb = RoundValueLessOne(ServingCarb,0.5,1)
        Else
            ServingCarb = 0.0
            roundCarb = 0.0
        End If
        'messagelist("serv Value TotCarb is: " & ServingCarb & " and rounded is: "&roundCarb)
        Dim SERVCTOTCARB As XmlNode
    	SERVCTOTCARB = xdTemplate.CreateElement("SERVCTOTCARB")
        xnRoot.AppendChild(SERVCTOTCARB)
        SERVCTOTCARB.InnerText = roundCarb.ToString()
        
        'Fiber
        Dim ServingFiber as object
        Dim fiberValue as object = ObjProperty("VALUE.TPALL", "" , "", "FIBER", "PARAM_CODE")
        if IsBlank(fiberValue) = 0 AndAlso Not IsDBNull(fiberValue) then
            ServingFiber = ( fiberValue * houseValue) / 100
        end if 
        
        Dim roundFiber as object 
        If Not IsBlank(ServingFiber) AndAlso Not IsDBNull(ServingFiber) AndAlso IsNumeric(ServingFiber) Then
            roundFiber = RoundValueLessOne(ServingFiber,0.5,1)
        Else
            ServingFiber = 0.0
            roundFiber = 0.0
        End If
        'messagelist("serv Value Fiber is: " & ServingFiber & " and rounded is: "& roundFiber)
        Dim SERVFIBER As XmlNode
    	SERVFIBER = xdTemplate.CreateElement("SERVFIBER")
        xnRoot.AppendChild(SERVFIBER)
        SERVFIBER.InnerText = roundFiber.ToString()
        
        'Sugar
        Dim ServingSug as object 
        Dim sugarValue as object = ObjProperty("VALUE.TPALL", "" , "", "SUGAR", "PARAM_CODE")
        if IsBlank(sugarValue) = 0 AndAlso Not IsDBNull(sugarValue) then
            ServingSug= (sugarValue * houseValue) / 100
        end if 
        
        Dim roundSugar as object 
        If Not IsBlank(ServingSug) AndAlso Not IsDBNull(ServingSug) AndAlso IsNumeric(ServingSug) Then
            roundSugar = RoundValueLessOne(ServingSug,0.5,1)
        Else
            ServingSug = 0.0
            roundSugar = 0.0
        End If
        'messagelist("serv Value Sugar is: " & ServingSug & " and rounded is: "& roundSugar)
        Dim SERVSUGAR As XmlNode
    	SERVSUGAR = xdTemplate.CreateElement("SERVSUGAR")
        xnRoot.AppendChild(SERVSUGAR)
        SERVSUGAR.InnerText = roundSugar.ToString()
        
        'Added Sugar
        Dim sugarAddValues as object = ObjProperty("VALUE.TPALL", "" , "", "ADDED_SUGAR_G", "PARAM_CODE")
        Dim ServingSugaradd as object
        if IsBlank(sugarAddValues) = 0 AndAlso Not IsDBNull(sugarAddValues) then
            ServingSugaradd = ( sugarAddValues * houseValue) / 100
        end if 
        
        Dim roundAddSugar as object 
        If Not IsBlank(ServingSugaradd) AndAlso Not IsDBNull(ServingSugaradd) AndAlso IsNumeric(ServingSugaradd) Then
            roundAddSugar = RoundValueLessOne(ServingSugaradd,0.5,1)
        Else
            ServingSugaradd = 0.0
            roundAddSugar = 0.0
        End If
        'messagelist("serv Value Added Sugar is: " & ServingSugaradd & " and rounded is: "& roundAddSugar)
        Dim SERVADDSUGAR As XmlNode
    	SERVADDSUGAR = xdTemplate.CreateElement("SERVADDSUGAR")
        xnRoot.AppendChild(SERVADDSUGAR)
        SERVADDSUGAR.InnerText = roundAddSugar.ToString()
        
        'added sugar without rounding for sample report
        Dim SAMPLESUGAR As XmlNode
    	SAMPLESUGAR = xdTemplate.CreateElement("SAMPLEADDSUGAR")
        xnRoot.AppendChild(SAMPLESUGAR)
        SAMPLESUGAR.InnerText = ServingSugaradd
        
        'SugarAlcol
        Dim ServingSUgAlc as object
        Dim sugarAlcolValue as object = ObjProperty("VALUE.TPALL", "" , "", "POLYOLS", "PARAM_CODE")
        if IsBlank(sugarAlcolValue) = 0 AndAlso Not IsDBNull(sugarAlcolValue) then
            ServingSUgAlc = ( sugarAlcolValue * houseValue) / 100
        end if 
        
        Dim roundAlcolSugar as object
        If Not IsBlank(ServingSUgAlc) AndAlso Not IsDBNull(ServingSUgAlc) AndAlso IsNumeric(ServingSUgAlc) Then
            roundAlcolSugar = RoundValueLessOne(ServingSUgAlc,0.5,1)
        Else
            ServingSUgAlc = 0.0
            roundAlcolSugar = 0.0
        End If
        'messagelist("serv Value Sugar Alcohol is: " & ServingSUgAlc & " and rounded is: "& roundAlcolSugar)
        Dim SERVALCOLSUGAR As XmlNode
    	SERVALCOLSUGAR = xdTemplate.CreateElement("SERVALCOLSUGAR")
        xnRoot.AppendChild(SERVALCOLSUGAR)
        SERVALCOLSUGAR.InnerText = roundAlcolSugar.ToString()
        
        'Protein
        Dim ServingPro as object
        Dim proteinValue as object = ObjProperty("VALUE.TPALL", "" , "", "PROTEIN", "PARAM_CODE")
        if IsBlank(proteinValue) = 0 AndAlso Not IsDBNull(proteinValue) then
            ServingPro = ( proteinValue * houseValue ) / 100
        end if 
        
        Dim roundPro as object 
        If Not IsBlank(ServingPro) AndAlso Not IsDBNull(ServingPro) AndAlso IsNumeric(ServingPro) Then
            roundPro = RoundValueLessOne(ServingPro,0.5,1)
        Else
            ServingPro = 0.0
            roundPro = 0.0
        End If
        'messagelist("serv Value Protein is: " & ServingPro & " and rounded is: "& roundPro)
        Dim SERVPROTEIN As XmlNode
    	SERVPROTEIN = xdTemplate.CreateElement("SERVPROTEIN")
        xnRoot.AppendChild(SERVPROTEIN)
        SERVPROTEIN.InnerText = roundPro.ToString()
        
        'Vitamin D
        Dim ServingVitD as object
        Dim vitDMcg as object = ObjProperty("VALUE.TPALL","" , "", "VIT_D_MCG" , "PARAM_CODE")
        if Not IsBlank(vitDMcg) AndAlso Not IsDBNull(vitDMcg) AndAlso IsNumeric(vitDMcg) Then
            ServingVitD = (vitDMcg * houseValue) / 100
        end If
            
        Dim roundVitD as object 
        If Not IsBlank(ServingVitD) AndAlso Not IsDBNull(ServingVitD) AndAlso IsNumeric(ServingVitD) Then
            roundVitD = RoundingVitaminMineralAmount(ServingVitD)
        Else
            ServingVitD = 0.0
            roundVitD = 0.0
        End If
        'messagelist("serv Value Vit-D is: " & ServingVitD & " and rounded is: "& roundVitD)
        Dim SERVVITD As XmlNode
    	SERVVITD = xdTemplate.CreateElement("SERVVITD")
        xnRoot.AppendChild(SERVVITD)
        SERVVITD.InnerText = roundVitD.ToString()
        
        'Calcuim
        Dim ServingCal as object
        Dim calcValue as object = ObjProperty("VALUE.TPALL", "" , "", "CALCIUM", "PARAM_CODE")
        if Not IsBlank(calcValue) AndAlso Not IsDBNull(calcValue) AndAlso IsNumeric(calcValue) Then
            ServingCal = ( calcValue * houseValue) / 100
        end if 
        
        Dim roundCal as object
        If Not IsBlank(ServingCal) AndAlso Not IsDBNull(ServingCal) AndAlso IsNumeric(ServingCal) Then
            roundCal = RoundingVitaminMineralAmount(ServingCal)
        Else
            ServingCal = 0.0
            roundCal = 0.0
        End If
        'messagelist("serv Value Calcium is: " & ServingCal & " and rounded is: "& roundCal)
        Dim SERVCALCIUM As XmlNode
    	SERVCALCIUM = xdTemplate.CreateElement("SERVCALCIUM")
        xnRoot.AppendChild(SERVCALCIUM)
        SERVCALCIUM.InnerText = roundCal.ToString()
        
        'Iron
        Dim ServingIron as object 
        Dim ironvalue as object = ObjProperty("VALUE.TPALL", "" , "", "IRON", "PARAM_CODE")
        if Not IsBlank(ironvalue) AndAlso Not IsDBNull(ironvalue) AndAlso IsNumeric(ironvalue) Then
            ServingIron = ( ironvalue * houseValue) / 100
        end if 
        
        Dim roundIron as object
        If Not IsBlank(ServingIron) AndAlso Not IsDBNull(ServingIron) AndAlso IsNumeric(ServingIron) Then
            roundIron = RoundingVitaminMineralAmount(ServingIron)
        Else
            ServingIron = 0.0
            roundIron = 0.0
        End If
        'messagelist("serv Value Iron is: " & ServingIron & " and rounded is: "& roundIron)
        Dim SERVIRON As XmlNode
    	SERVIRON = xdTemplate.CreateElement("SERVIRON")
        xnRoot.AppendChild(SERVIRON)
        SERVIRON.InnerText = roundIron.ToString()
        
        'Potassium
        Dim ServingPot as object
        Dim potasValue as object = ObjProperty("VALUE.TPALL", "" , "", "POTASSIUM" , "PARAM_CODE")
        If Not IsBlank(potasValue) AndAlso Not IsDBNull(potasValue) AndAlso IsNumeric(potasValue) Then
            ServingPot = (potasValue * houseValue) / 100
        end if 
        
        Dim roundPotas as object
        If Not IsBlank(ServingPot) AndAlso Not IsDBNull(ServingPot) AndAlso IsNumeric(ServingPot) Then
            roundPotas = RoundingVitaminMineralAmount(ServingPot)
        Else
            ServingPot = 0.0
            roundPotas = 0.0
        End If
        'messagelist("serv Value Potas is: " & ServingPot & " and rounded is: "& roundPotas)
        Dim SERVPOTASSIUM As XmlNode
    	SERVPOTASSIUM = xdTemplate.CreateElement("SERVPOTASSIUM")
        xnRoot.AppendChild(SERVPOTASSIUM)
        SERVPOTASSIUM.InnerText = roundPotas.ToString()
        
        
        'parameters of percentage of daily value
        Dim nutritionParam() as string = {"FAT","FASAT_G","CHOLE_MG","CARB_TOTAL_AMER","FIBER","SUGAR_ADD","PROTEIN","VIT_D_MCG","CALCIUM","IRON","POTASSIUM"} 
        
        ' Define the fixed Current Daily Values in an array for calculating %Daily values
        Dim dailyValues() As Double = {78.0, 20.0, 300.0, 275.0, 28.0, 50.0, 50.0, 20.0, 1300.0, 18.0, 4700.0}
        
        Dim perservingValues() As Double = {ServingFat,ServingFatSat,ServingChole,ServingCarb,ServingFiber,ServingSugaradd,ServingPro,ServingVitD,ServingCal,ServingIron,ServingPot}
        
        'calculating percentage dailyvalue
        Dim DailyValueNodes As XmlNode
        Dim percentDailyValue(perservingValues.Length - 1) As Double 
        For i As Integer = 0 To perservingValues.Length - 1
            if nutritionParam(i) = "VIT_D_MCG" or nutritionParam(i) = "CALCIUM" or nutritionParam(i) = "IRON" or nutritionParam(i) = "POTASSIUM" then
                percentDailyValue(i) = RoundPDV((perservingValues(i) / dailyValues(i)) * 100)
                'nodes of dailyvalues
                DailyValueNodes = xdTemplate.CreateElement("DAILY"&nutritionParam(i))
                xnRoot.AppendChild(DailyValueNodes)
                DailyValueNodes.InnerText = percentDailyValue(i).ToString()
                
            else
                
                percentDailyValue(i) = Math.Truncate((perservingValues(i) / dailyValues(i)) * 100) 
                'nodes of dailyvalues
                DailyValueNodes = xdTemplate.CreateElement("DAILY"&nutritionParam(i))
                xnRoot.AppendChild(DailyValueNodes)
                DailyValueNodes.InnerText = percentDailyValue(i).ToString()
            end if
            'messagelist(nutritionParam(i) & "  daily value% is  " & percentDailyValue(i) )
        Next
        
        'for daily &serving value of sodium(because it starts with 15 and cannot be the name of a xml node)
    	Dim ServingSodiumoptiva As Object = ObjProperty("VALUE.TPALL", "", "", "15_SODIUM_100", "PARAM_CODE")
    	Dim ServingSodium as Object
    	
    	If Not IsBlank(ServingSodiumoptiva) AndAlso Not IsDBNull(ServingSodiumoptiva) AndAlso IsNumeric(ServingSodiumoptiva) Then 
    	    ServingSodium = (ServingSodiumoptiva * houseValue) / 100
    	end if 
    	
    	Dim roundSodium as Object
    	If Not IsBlank(ServingSodium) AndAlso Not IsDBNull(ServingSodium) AndAlso IsNumeric(ServingSodium) Then
    	    roundSodium = RoundValueGreaterOne(ServingSodium, 5 , 140 , 5 , 10)
        else
            ServingSodium = 0.0
            roundSodium = 0.0
    	end if
        'messagelist("serv Value Sodium is: " & ServingSodium & " and rounded is: "& roundSodium)
        Dim SERVSODIUM As XmlNode
    	SERVSODIUM = xdTemplate.CreateElement("SERVSODIUM")
        xnRoot.AppendChild(SERVSODIUM)
        SERVSODIUM.InnerText = roundSodium.ToString()
    	
    	'Dim recalccurrentvalue as object = (2300 * ServingSodium)/ ServingSodiumoptiva
    	Dim sodiumpercentDaily as Integer = CInt(Math.Truncate((ServingSodium / 2300) * 100))
    	'messagelist( "final percentdaily sodium is: "&  sodiumpercentDaily)
    	Dim DAILYSODIUM As XmlNode
    	DAILYSODIUM = xdTemplate.CreateElement("DAILYSODIUM")
        xnRoot.AppendChild(DAILYSODIUM)
        DAILYSODIUM.InnerText = sodiumpercentDaily.ToString()
        'messagelist("Sodium daily Value is:  " & sodiumpercentDaily)
        
        '========================================================================================
        ' END - Add Samples data to the report XML template for NFP
        '========================================================================================
        
    
    Catch ex as exception
    	messagelist("Report Generation Error:  ", ex.message)
    End try

    return 111

End Function

Function RoundValueHalfOne(paramvalue As Double, threshold1 As Double, threshold2 As Double) As object 
    If paramvalue < threshold1 Then
        Return 0
    ElseIf paramvalue < threshold2 Then
        Return Math.Round(paramvalue * 2, MidpointRounding.AwayFromZero) / 2
    Else
        Return Math.Round(paramvalue, MidpointRounding.AwayFromZero)
    End If
End Function

Function RoundValueGreaterOne(paramvalue As Object, threshold1 As Double, threshold2 As Double, rounded1 As Double, rounded2 As Double) As Object
    Dim value As Double
    
    ' Check if paramvalue is nothing or if it is an empty string
    If paramvalue Is Nothing OrElse String.IsNullOrWhiteSpace(paramvalue.ToString()) Then
        Return 0
    End If

    ' Attempt to convert paramvalue to Double
    If Not Double.TryParse(paramvalue.ToString(), value) Then
        Return 0
    End If

    ' Proceed with rounding logic
    If value < threshold1 Then
        Return 0
    ElseIf value < threshold2 Then
        Return Math.Round(value / rounded1, MidpointRounding.AwayFromZero) * rounded1
    Else
        Return Math.Round(value / rounded2, MidpointRounding.AwayFromZero) * rounded2
    End If
End Function



Function RoundValueLessOne(paramvalue As Double, threshold1 As Double, threshold2 As Double) As object
    If paramvalue < threshold1 Then
        Return 0
    ElseIf paramvalue < threshold2 Then
        Return "<1"
    Else
        Return Math.Round(paramvalue, MidpointRounding.AwayFromZero)
    End If
End Function

Function RoundingVitaminMineralAmount(amount As Double) As object
    If amount = Math.Floor(amount) Then
        Return amount.ToString("0")
    Else
        Return amount.ToString("0.##")
    End If
End Function

Function RoundPDV(value As Double) As object
    If value < 2 Then
        Return 0
    ElseIf value <= 10 Then
        Return Math.Round(value / 2) * 2
    ElseIf value <= 50 Then
        Return Math.Round(value / 5) * 5
    Else
        Return Math.Round(value / 10) * 10
    End If
End Function

End Class
