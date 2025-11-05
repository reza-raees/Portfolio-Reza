'Creation of a Function to Automaticly fill the Extension Table of the item by an appropriate Core team

Sub GetRelatedCoreTeam(ByVal subbrand As String , ByRef team() As String , ByRef depart() As String)

    Select Case UCase(subbrand)
        Case "MAGIC CARE"
            team = {"Priska Stich", "Marilu Pötter", "Laura Thoma", "Kathleen Mai", "Antje Trieder"}
            depart = {"R&D", "Brand Management", "Product Management", "R&D", "R&D"}

        Case "AQUA INTEN"
            team = {"Nicole Delitz", "Suad Jahic", "Metin Önul", "Antje Trieder", "Seok Won Jung"}
            depart = {"Product Management", "R&D", "Brand Management", "R&D", "R&D"}

        Case "AHUHU"
            team = {"Andrea Seidl", "Carola Greess", "Jessica Rempel", "Karin Enthaler", "Kristine Müller", "Lisa-Marie Böhm", "Lukas Fischer", "Manuela Mischke", "Nicole Delitz", "Timur Scharhag"}
            depart = {"", "Brand Management", "R&D", "R&D", "", "R&D", "", "", "Product Management", ""}

        Case "MAGIC FINI"
            team = {"Martina Berger", "Stefanie Huber", "Laura Hoppe", "Marilu Pötter", "Vanessa Bahnsen", "Angelina Lichtenstern", "Jessica Rempel", "Priyanka Mahajan", "Anja Reinhardt", "Antonia Mariß"}
            depart = {"Product Management", "Product Management", "Product Management", "Brand Management", "R&D", "Brand Management", "R&D", "R&D", "", ""}

        Case "COLLA LIFT"
            team = {"Nicole Delitz", "Angelina Lichtenstern", "Kathleen Mai", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "CLEAR SKIN"
            team = {"Laura Thoma", "Katharina Franzske", "Kathleen Mai", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "OCEAN MINE"
            team = {"Katharina Franzske", "Laura Hoppe", "Antje Trieder"}
            depart = {"Brand Management", "Product Management", "R&D"}

        Case "RES PREM50"
            team = {"Rüdiger Knapp", "Astrid Sonnenstatter", "Antje Trieder", "Kathleen Mai"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "RETINOLINT"
            team = {"Nicole Delitz", "Astrid Sonnenstatter", "Kathleen Mai", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "VINO GOLD"
            team = {"Rüdiger Knapp", "Astrid Sonnenstatter", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D"}

        Case "AGE EFFECT"
            team = {"Rüdiger Knapp", "Carola Greess", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D"}

        Case "VINOLIFT"
            team = {"Rüdiger Knapp", "Astrid Sonnenstatter", "Kathleen Mai", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "VITAMIN C"
            team = {"Rüdiger Knapp", "Metin Önul", "Marilu Pötter", "Suad Jahic", "Seok Won Jung", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "Brand Management", "R&D", "R&D", "R&D"}

        Case "VITAMIN E"
            team = {"Nicole Delitz", "Priska Stich", "Antje Trieder", "Astrid Sonnenstatter"}
            depart = {"Product Management", "R&D", "R&D", "Brand Management"}

        Case "FINE FRAGR"
            team = {"Laura Hoppe", "Katharina Franzske", "Seok Won Jung", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "BATH&BODY"
            team = {"Laura Hoppe", "Katharina Franzske", "Seok Won Jung", "Antje Trieder"}
            depart = {"Product Management", "Brand Management", "R&D", "R&D"}

        Case "SUN"
            team = {"Carola Greess", "Katharina Franzske", "Antje Trieder", "Nicole Delitz"}
            depart = {"Brand Management", "Brand Management", "R&D", "Product Management"}

        Case Else
            team = {}
            depart = {}
    End Select

End Sub




'Calling the Function in the Save script

'===========================================================================
'Begin  -   Reza.     Updating the Core Team Extension table 
'===========================================================================
Dim sBrand as String = ObjProperty("BRAND","","")
Dim subbrand as string = ObjProperty("SUBBRAND","","")

if isblank(subbrand) = 0 then
	if sBrand = "2" then subbrand = "AHUHU"
	
	Dim team() As String
	Dim dept() As String
	genfun.GetRelatedCoreTeam(subbrand, team, dept)
	for i as integer = 0 to team.length -1
		Dim newDM0Row As DataRow = GetNewRow("", "", "MATRIX.V\DMFORMULA4")
		newDM0Row("FIELD1") = team(i)
		newDM0Row("FIELD2") = dept(i)
		CommitNewRow("", "", "MATRIX.V\DMFORMULA4", newDM0Row)
	next
end if
'===========================================================================
'End    -   Reza.     Updating the Core Team Extension table 
'===========================================================================




'Automating to Fullfill another Extension table*************************************************************************************************************** 

'==============================================================================================
'Begin  -   Reza.     22/10/2025      Updating the Lab Trial Extension tables for Bulk Formulas 
'==============================================================================================
Dim sClass As String = ObjProperty("CLASS","","")
if sClass = "BULK_FORMULA" or sClass = "DEV_BULK_FORMULA" then
	
	'Dim sDate      As String = "we will add in the future"    
	Dim Appearance As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_APPEARANCE" ,1 )
	Dim Colour     As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_COLOUR" ,1 )
	Dim Odour      As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_ODOUR" ,1 )
	'Dim Start_pH   As String = ObjProperty("VALUE.TPALL","","","SO_PRE_REL_PH_VALUE","PARAM_CODE")
	Dim Ph         As String = ObjProperty("VALUE.TPALL","","","SO_PH_VALUE","PARAM_CODE")
	Dim Density    As String = ObjProperty("VALUE.TPALL","","","SO_BULK_DENSITY","PARAM_CODE")
	Dim Viscosity  As String = ObjProperty("VALUE.TPALL","","","SO_VISCOSITY","PARAM_CODE")
	Dim Centrifuge As String = EnumLabel("C_SO_CENTRIFUGE",ObjProperty("VALUE.TPALL","","","SO_CENTRIFUGE","PARAM_CODE"),"EN-US")
	Dim Melting    As String = ObjProperty("VALUE.TPALL","","","SO_MELTING_POINT","PARAM_CODE")
	Dim PreAppearance As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_PRE_REL_APPEAR" ,1 )
	Dim PreColour     As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_PRE_REL_COLOUR" ,1 )
	Dim PreOdour      As String = ObjProperty("DOCTEXT.DOC.A", "", "", "SO_PRE_REL_ODOUR" ,1 )
	Dim PrePh         As String = ObjProperty("VALUE.TPALL","","","SO_PRE_REL_PH_VALUE","PARAM_CODE")
	Dim PreDensity    As String = ObjProperty("VALUE.TPALL","","","SO_PRE_REL_DENSITY","PARAM_CODE")
	Dim PreViscosity  As String = ObjProperty("VALUE.TPALL","","","SO_PRE_REL_VISCOSITY","PARAM_CODE")
	Dim PreCentrifuge As String = EnumLabel("C_SO_CENTRIFUGE",ObjProperty("VALUE.TPALL","","","SO_PRE_REL_CENTRIF","PARAM_CODE"),"EN-US")
	
	
	'ObjPropertySet(sDate, 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "1" , "FIELD1" )
	ObjPropertySet(Appearance, 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "2" , "FIELD1" )
	ObjPropertySet(Colour    , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "3" , "FIELD1" )
	ObjPropertySet(Odour     , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "4" , "FIELD1" )
	'ObjPropertySet(Start_pH  , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "5" , "FIELD1" )
	ObjPropertySet(Ph        , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "6" , "FIELD1" )
	ObjPropertySet(Density   , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "7" , "FIELD1" )
	ObjPropertySet(Viscosity , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "8" , "FIELD1" )
	ObjPropertySet(Centrifuge, 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "9" , "FIELD1" )
	ObjPropertySet(Melting   , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "10" , "FIELD1")
	ObjPropertySet(PreAppearance, 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "11" , "FIELD1" )
	ObjPropertySet(PreColour    , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "12" , "FIELD1" )
	ObjPropertySet(PreOdour     , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "13" , "FIELD1" )
	ObjPropertySet(PrePh        , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "14" , "FIELD1" )
	ObjPropertySet(PreDensity   , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "15" , "FIELD1" )
	ObjPropertySet(PreViscosity , 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "16" , "FIELD1" )
	ObjPropertySet(PreCentrifuge, 1 , "FIELD2.MATRIX.V\DMFORMULA1" , "" , "" , "17" , "FIELD1" )


End if
'==============================================================================================
'End    -   Reza.     22/10/2025      Updating the Lab Trial Extension tables for Bulk Formulas 
'==============================================================================================