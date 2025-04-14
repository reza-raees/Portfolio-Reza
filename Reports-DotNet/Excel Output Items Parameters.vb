' Purpose:  Print Formula's Nutrituin Panel Facts

' AUTHOR:	Reza     

' NOTE: 	This file includes only the Developments for creating the Items Parameters Information report using 'DO_ITEM_EXCEL_REPORT' table from database.
' 			It does not contain full Interface Configuration or system templates generated.Option Strict Off


Imports System
Imports System.Diagnostics
Imports System.Collections
Imports Microsoft.VisualBasic
Imports System.data
Imports System.Drawing
Imports OfficeOpenXml.Drawing
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Net
Imports System.IO
Imports ICPVar = Infor.DocumentManagement.ICP
Imports System.Collections.Generic
Imports System.linq
Imports System.Math
Imports System.XML

Class ActionScript
	Inherits FcProcFuncSetEventWF
	
	'for retreiving the status label
	dim gf as GENERALFUNCTIONS = New GENERALFUNCTIONS(Me) 
	
	Function wf_start() As Long
		Return 1
	End Function
	
	Function wf_complete() As Long
	Try
	
		'Create Excel Spreadsheet File
		Dim pkg as New ExcelPackage()
		
		'Add a worksheet to the Excel Spreadsheet
		Dim worksheet As ExcelWorksheet = pkg.workbook.Worksheets.Add("Report Class Recognition")
		Dim rowIndex As Integer = 0
		
		'Define the default font and size
		worksheet.Cells.Style.font.Size= 10
		'worksheet.Cells.Style.Font.Name = "Calibri" '15 - LINE37
		Dim ColumnWidth() As Double = {15,15,20,35,35,50,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,20,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,20,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15,15}
		
		For I as Integer = 1 To columnWidth.Length
			Worksheet.Column(i).Width = columnWidth(i - 1)
			Worksheet.Column(i).Style.WrapText = True
			Worksheet.Column(i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
		Next i
		
		'Worksheet.Cells(1 , 1).Value = "Report Damiano's Item"
		'Worksheet.Cells(1 , 1).Style.Font.Bold = True
		'Worksheet.Cells(1 , 1 , 1 ,2).Merge = True	
		
		' Set the first title
        Worksheet.Cells(1, 1).Value = "DATA PRINCIPALE"
        Worksheet.Cells(1, 1).Style.Font.Bold = True
        Worksheet.Cells(1, 1, 1, 21).Merge = True
        Worksheet.Cells(1, 1, 1, 21).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        
        Worksheet.Cells(1, 22).Value = "ANAGRAFICA"
        Worksheet.Cells(1, 22).Style.Font.Bold = True
        Worksheet.Cells(1, 22, 1, 64).Merge = True
        Worksheet.Cells(1, 22, 1, 64).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        
        Worksheet.Cells(1, 65).Value = "CERTIFICAZIONE"
        Worksheet.Cells(1, 65).Style.Font.Bold = True
        Worksheet.Cells(1, 65, 1, 84).Merge = True
        Worksheet.Cells(1, 65, 1, 84).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        
		Worksheet.Cells(1, 85).Value = "QUALITA"
        Worksheet.Cells(1, 85).Style.Font.Bold = True
        Worksheet.Cells(1, 85, 1, 105).Merge = True
        Worksheet.Cells(1, 85, 1, 105).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        
        Worksheet.Cells(1, 106).Value = "REPORT_BIO"
        Worksheet.Cells(1, 106).Style.Font.Bold = True
        Worksheet.Cells(1, 106, 1, 108).Merge = True
        Worksheet.Cells(1, 106, 1, 108).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
        
        Worksheet.Cells(1, 109).Value = "NO CATEGORY"
        Worksheet.Cells(1, 109).Style.Font.Bold = True
        Worksheet.Cells(1, 109, 1, 119).Merge = True
        Worksheet.Cells(1, 109, 1, 119).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
		
		'Table Header
		worksheet.Row(2).Height = 40
	    
		'DATA PRINCIPALI CATAGORY
		worksheet.Cells("A2").Value = "ITEM_CODE"
        worksheet.Cells("B2").Value = "TYPE"
        worksheet.Cells("C2").Value = "DO_DESC_BREVE"
        worksheet.Cells("D2").Value = "DESC_DOCUMENTI"       
        worksheet.Cells("E2").Value = "DESC_ESTESA"            
        worksheet.Cells("F2").Value = "CODICE_CONTROLLO"
        worksheet.Cells("G2").Value = "CREATION_DATE"
        worksheet.Cells("H2").Value = "CREATED_BY"
        worksheet.Cells("I2").Value = "APPROVAZIONE"
        worksheet.Cells("J2").Value = "DO_DESC_DOC_FLAG"
        worksheet.Cells("K2").Value = "DO_DESC_EST_FLAG"
        worksheet.Cells("L2").Value = "CODICE_SAP_BIO"
        worksheet.Cells("M2").Value = "DO_TESTO_ODA"
        worksheet.Cells("N2").Value = "DESCRIZIONE_COMMERCIALE"
        worksheet.Cells("O2").Value = "DATA_MODIFICA"
        worksheet.Cells("P2").Value = "MODIFY_BY"
        worksheet.Cells("Q2").Value = "STATUS"
        worksheet.Cells("R2").Value = "CLASS"
        worksheet.Cells("S2").Value = "UOM_CODE"
        worksheet.Cells("T2").Value = "DO_DESC_BREVE_FLAG"
        worksheet.Cells("U2").Value = "FORMULA_ETICHETTA_ALIMENTARE"
        
        'ANAGRAFICA CATAGORY
        worksheet.Cells("V2").Value = "DO_AMB_SAP"
        worksheet.Cells("W2").Value = "DO_BP"
        worksheet.Cells("X2").Value = "DO_CARATTERISTICA"
        worksheet.Cells("Y2").Value = "DO_CAT_MERCE"
        worksheet.Cells("Z2").Value = "DO_CERT_IMB"
        worksheet.Cells("AA2").Value = "DO_CICLO"
        worksheet.Cells("AB2").Value = "DO_COMPOSIZIONE"
        worksheet.Cells("AC2").Value = "DO_CONFEZIONAMENTO"
        worksheet.Cells("AD2").Value = "DO_CONTO_PROP_TERZI"
        worksheet.Cells("AE2").Value = "DO_DES_PROD"
        worksheet.Cells("AF2").Value = "DO_DESTINAZIONE"
        worksheet.Cells("AG2").Value = "DO_DIVISIONE"
        worksheet.Cells("AH2").Value = "DO_EAN_CODE"
        worksheet.Cells("AI2").Value = "DO_FAM_PRD_WMS"
        worksheet.Cells("AJ2").Value = "DO_FORMATO_PACK"
        worksheet.Cells("AK2").Value = "DO_GEST_LOTTO_CLI"
        worksheet.Cells("AL2").Value = "DO_GEST_LOTTO_WMS"
        worksheet.Cells("AM2").Value = "DO_GG_MEDI_CQ"
        worksheet.Cells("AN2").Value = "DO_GRP_CONTAB"
        worksheet.Cells("AO2").Value = "DO_GRP_PRD"
        worksheet.Cells("AP2").Value = "DO_IMBALLO_SML"
        worksheet.Cells("AQ2").Value = "DO_INGREDIENTI"
        worksheet.Cells("AR2").Value = "DO_LAVORAZIONE"
        worksheet.Cells("AS2").Value = "DO_LINEA_BUSINESS"
        worksheet.Cells("AT2").Value = "DO_LINGUA"
        worksheet.Cells("AU2").Value = "DO_LINGUA_ETICHETTA"
        worksheet.Cells("AV2").Value = "DO_MACRO_CERT"
        worksheet.Cells("AW2").Value = "DO_MARCHIO"
        worksheet.Cells("AX2").Value = "DO_MERCE_SELEZIONATA"
        worksheet.Cells("AY2").Value = "DO_MONO_MULTIORIGINE"
        worksheet.Cells("AZ2").Value = "DO_ORIG_DOGANALE"
        worksheet.Cells("BA2").Value = "DO_PAESE_ORIGINE"
        worksheet.Cells("BB2").Value = "DO_RADICE"
        worksheet.Cells("BC2").Value = "DO_REGIONE_ORIGINE"
        worksheet.Cells("BD2").Value = "DO_RICETTA"
        worksheet.Cells("BE2").Value = "DO_SETT_MERCE"
        worksheet.Cells("BF2").Value = "DO_SINGLE_MULTIBLEND"
        worksheet.Cells("BG2").Value = "DO_STATUS"
        worksheet.Cells("BH2").Value = "DO_TIPO"
        worksheet.Cells("BI2").Value = "DO_TIPO_IMBALLO"
        worksheet.Cells("BJ2").Value = "DO_TIPO_PRD"
        worksheet.Cells("BK2").Value = "DO_TIPO_SML"
        worksheet.Cells("BL2").Value = "DO_VARIANTE"
        
        'CERTIFICATION CATAGORY
        worksheet.Cells("BM2").Value = "DO_FAIR_TRADE"
        worksheet.Cells("BN2").Value = "DO_FFL"
        worksheet.Cells("BO2").Value = "DO_FILIERA_CONV"
        worksheet.Cells("BP2").Value = "DO_FILIERA_CONV_CERT"
        worksheet.Cells("BQ2").Value = "DO_HALAL"
        worksheet.Cells("BR2").Value = "DO_IBD"
        worksheet.Cells("BS2").Value = "DO_JAS"
        worksheet.Cells("BT2").Value = "DO_KOSHER"
        worksheet.Cells("BU2").Value = "DO_NATURLAND"
        worksheet.Cells("BV2").Value = "DO_NOP"
        worksheet.Cells("BW2").Value = "DO_OIA"
        worksheet.Cells("BX2").Value = "DO_PVB"
        worksheet.Cells("BY2").Value = "DO_REGEN_ORGANIC"
        worksheet.Cells("BZ2").Value = "DO_SPIGA_BARRATA"
        worksheet.Cells("CA2").Value = "DO_VEGAN"
        worksheet.Cells("CB2").Value = "DO_BIO"
        worksheet.Cells("CC2").Value = "DO_BIO_SUISSE"
        worksheet.Cells("CD2").Value = "DO_CONVENZIONALE"
        worksheet.Cells("CE2").Value = "DO_CMO"
        worksheet.Cells("CF2").Value = "DO_DEMETER"
        
        'QUALITA CATAGORY
        worksheet.Cells("CG2").Value = "DO_ALLERGENI"
        worksheet.Cells("CH2").Value = "DO_BIOLOGICO"
        worksheet.Cells("CI2").Value = "DO_CAR_COLORE_A"
        worksheet.Cells("CJ2").Value = "DO_CAR_COLORE_B"
        worksheet.Cells("CK2").Value = "DO_CAR_COLORE_L"
        worksheet.Cells("CL2").Value = "DO_CODICE_CONAI"
        worksheet.Cells("CM2").Value = "DO_DEN_CONV_UDB"
        worksheet.Cells("CN2").Value = "DO_FREDDO"
        worksheet.Cells("CO2").Value = "DO_LOTTO_PROPR_CLI"
        worksheet.Cells("CP2").Value = "DO_MATERIALE_IMBALLO"
        worksheet.Cells("CQ2").Value = "DO_NUM_CONV_UDB"
        worksheet.Cells("CR2").Value = "DO_PESO_LORDO"
        worksheet.Cells("CS2").Value = "DO_PESO_NETTO"
        worksheet.Cells("CT2").Value = "DO_SECONDO_CONTROLLO"
        worksheet.Cells("CU2").Value = "DO_SIZE"
        worksheet.Cells("CV2").Value = "DO_SLUNGA_COMPLIANT"
        worksheet.Cells("CW2").Value = "DO_SMT"
        worksheet.Cells("CX2").Value = "DO_TEMP_TOSTATURA"
        worksheet.Cells("CY2").Value = "DO_TRACCIABILE"
        worksheet.Cells("CZ2").Value = "DO_UDM_ALT_ISO"
        worksheet.Cells("DA2").Value = "SHELF_LIFE"
        
        'REPORT BIO CATAGORY
        worksheet.Cells("DB2").Value = "DO_FAM_REPORT_BIO_SP"
        worksheet.Cells("DC2").Value = "DO_IN_REPORT_BIO"
        worksheet.Cells("DD2").Value = "DO_IN_REPORT_BIO_SPP"

        'BLANK(NO) CATAGORY
        worksheet.Cells("DE2").Value = "DO_CONV_CERTIFICATO"
        worksheet.Cells("DF2").Value = "DO_DAMIANO_SUPER_DOP"
        worksheet.Cells("DG2").Value = "DO_DOP"
        worksheet.Cells("DH2").Value = "DO_DOP_BIO"
        worksheet.Cells("DI2").Value = "DO_FILIERA_BIO_GOLD"
        worksheet.Cells("DJ2").Value = "DO_FILIERA_BIO_PREM"
        worksheet.Cells("DK2").Value = "DO_FILIERA_BIO_STAND"
        worksheet.Cells("DL2").Value = "DO_IGP"
        worksheet.Cells("DM2").Value = "DO_IGP_BIO"
        worksheet.Cells("DN2").Value = "DO_NESSUNA_CERT_EST"
        worksheet.Cells("DO2").Value = "DO_NESSUNA_CERT_INT"

        'setting of the costumization   
		worksheet.Cells("A2:DO2").Style.Font.Bold = True
		worksheet.Cells("A2:DO2").Style.Border.Left.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A2:DO2").Style.Border.Top.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A2:DO2").Style.Border.Right.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A2:DO2").Style.Border.Bottom.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A1:DO2").Style.Fill.PatternType = ExcelFillStyle.Solid
		worksheet.Cells("A1:U1").Style.Fill.BackgroundColor.SetColor(Color.LightGreen)
		worksheet.Cells("V1:BL1").Style.Fill.BackgroundColor.SetColor(Color.LightYellow)
		worksheet.Cells("BM1:CF1").Style.Fill.BackgroundColor.SetColor(Color.LightBlue)
		worksheet.Cells("CG1:DA1").Style.Fill.BackgroundColor.SetColor(Color.red)
		worksheet.Cells("DB1:DD1").Style.Fill.BackgroundColor.SetColor(Color.Yellow)
		worksheet.Cells("DE1:DO1").Style.Fill.BackgroundColor.SetColor(Color.Green)
		worksheet.Cells("A2:U2").Style.Fill.BackgroundColor.SetColor(Color.LightGreen)
		worksheet.Cells("V2:BL2").Style.Fill.BackgroundColor.SetColor(Color.LightYellow)
		worksheet.Cells("BM2:CF2").Style.Fill.BackgroundColor.SetColor(Color.LightBlue)
		worksheet.Cells("CG2:DA2").Style.Fill.BackgroundColor.SetColor(Color.red)
		worksheet.Cells("DB2:DD2").Style.Fill.BackgroundColor.SetColor(Color.Yellow)
		worksheet.Cells("DE2:DO2").Style.Fill.BackgroundColor.SetColor(Color.Green)
		
		'Getting Input
		Dim sClass as String = WIPParamGet("CLASS_REC")
		
		Dim pDO_FAIR_TRADE as String = WIPParamGet("DO_FAIR_TRADE")
		Dim pDO_GRP_PRD as String = WIPParamGet("DO_GRP_PRD")
		Dim pDO_LINEA_BUSINESS as String = WIPParamGet("DO_LINEA_BUSINESS")
		Dim pDO_MACRO_CERT as String = WIPParamGet("DO_MACRO_CERT")
		Dim pDO_MARCHIO as String = WIPParamGet("DO_MARCHIO")
		Dim pDO_TIPO_PRD as String = WIPParamGet("DO_TIPO_PRD")
		
		Dim otptClass() as string = sClass.Split(";")
		Dim iStartRow As Integer = 2
		
		'Filtering the data regarding the Class
		iStartRow += 1
		Dim itemTable as DataTable = TABLELOOKUPEX("DO_ITEM_EXCEL_REPORT","TABLELOOKUP",pDO_FAIR_TRADE,pDO_GRP_PRD,pDO_LINEA_BUSINESS,pDO_MACRO_CERT,pDO_MARCHIO,pDO_TIPO_PRD)
		If Not itemTable Is Nothing andalso itemTable.rows.count > 0 then
			For Each filcls as string in otptClass
				Dim ibresult() as  Datarow = itemTable.select("CLASS = '"+ filcls +"'")
				For each drItem as datarow in ibresult
				    
				    ' DATA PRINCIPALI CATEGORY
                    worksheet.Cells(iStartRow, 1).Value = CStr(drItem("ITEM_CODE"))
                    worksheet.Cells(iStartRow, 2).Value = GetValueOrEnumLabel("C_COMPONENT_IND", drItem("COMPONENT_IND")) 
                    worksheet.Cells(iStartRow, 3).Value = drItem("DO_DESC_BREVE")
                    worksheet.Cells(iStartRow, 4).Value = drItem("DESC_DOCUMENTI")
                    worksheet.Cells(iStartRow, 5).Value = drItem("DESC_ESTESA")
                    worksheet.Cells(iStartRow, 6).Value = drItem("CODICE_CONTROLLO")
                    worksheet.Cells(iStartRow, 7).Value = drItem("CREATION_DATE")
                    worksheet.Cells(iStartRow, 8).Value = drItem("CREATED_BY")
                    worksheet.Cells(iStartRow, 9).Value = drItem("APPROVAL_CODE")
                    
                    if drItem("DO_DESC_DOC_FLAG").ToString() = "0" or drItem("DO_DESC_DOC_FLAG").ToString() = "" Then
                        worksheet.Cells(iStartRow, 10).Value = "FALSO"
                    elseif drItem("DO_DESC_DOC_FLAG").ToString() = "1"
                        worksheet.Cells(iStartRow, 10).Value = " "
                    end if
                    
                    if drItem("DO_DESC_EST_FLAG").ToString() = "0" or drItem("DO_DESC_EST_FLAG").ToString() = "" Then
                        worksheet.Cells(iStartRow, 11).Value = "FALSO"
                    elseif drItem("DO_DESC_EST_FLAG").ToString() = "1"
                        worksheet.Cells(iStartRow, 11).Value = " "
                    end if
                    
                    worksheet.Cells(iStartRow, 12).Value = drItem("CODICE_SAP_BIO")
                    worksheet.Cells(iStartRow, 13).Value = drItem("DO_TESTO_ODA")
                    worksheet.Cells(iStartRow, 14).Value = drItem("DESCRIZIONE_COMMERCIALE")
                    worksheet.Cells(iStartRow, 15).Value = drItem("MODIFY_DATE")
                    worksheet.Cells(iStartRow, 16).Value = drItem("MODIFY_BY")
                    worksheet.Cells(iStartRow, 17).Value = gf.GetStatusDesc("ITEM",drItem("STATUS_IND"))
                    worksheet.Cells(iStartRow, 18).Value = drItem("CLASS")
                    worksheet.Cells(iStartRow, 19).Value = drItem("UOM_CODE")
                    
                    if drItem("DO_DESC_BREVE_FLAG").ToString() = "0" or drItem("DO_DESC_BREVE_FLAG").ToString() = "" Then
                        worksheet.Cells(iStartRow, 20).Value = "FALSO"
                    elseif drItem("DO_DESC_BREVE_FLAG").ToString() = "1"
                        worksheet.Cells(iStartRow, 20).Value = " "
                    end if
                    
                    worksheet.Cells(iStartRow, 21).Value = drItem("FORMULA_CODE") + " \ " + drItem("VERSION") 
                    
                    ' ANAGRAFICA CATEGORY
                    worksheet.Cells(iStartRow, 22).Value = GetValueOrEnumLabel("C_AMB_SAP", drItem("DO_AMB_SAP"))
                    worksheet.Cells(iStartRow, 23).Value = GetValueOrEnumLabel("C_BP", drItem("DO_BP"))
                    worksheet.Cells(iStartRow, 24).Value = GetValueOrEnumLabel("C_CARATTERISTICA", drItem("DO_CARATTERISTICA"))
                    worksheet.Cells(iStartRow, 25).Value = GetValueOrEnumLabel("C_CAT_MERCE", drItem("DO_CAT_MERCE"))
                    worksheet.Cells(iStartRow, 26).Value = GetValueOrEnumLabel("C_CERT_IMB", drItem("DO_CERT_IMB"))
                    worksheet.Cells(iStartRow, 27).Value = GetValueOrEnumLabel("C_CICLO", drItem("C_CICLO"))
                    worksheet.Cells(iStartRow, 28).Value = GetValueOrEnumLabel("C_COMPOSIZIONE", drItem("DO_COMPOSIZIONE"))
                    worksheet.Cells(iStartRow, 29).Value = GetValueOrEnumLabel("C_CONFEZIONAMENTO", drItem("DO_CONFEZIONAMENTO"))
                    worksheet.Cells(iStartRow, 30).Value = GetValueOrEnumLabel("C_CONTO_PROP_TERZI", drItem("DO_CONTO_PROP_TERZI"))
                    worksheet.Cells(iStartRow, 31).Value = GetValueOrEnumLabel("C_DES_PROD", drItem("DO_DES_PROD"))
                    worksheet.Cells(iStartRow, 32).Value = GetValueOrEnumLabel("C_DESTINAZIONE", drItem("DO_DESTINAZIONE"))
                    worksheet.Cells(iStartRow, 33).Value = GetValueOrEnumLabel("C_DIVISIONE", drItem("DO_DIVISIONE"))
                    worksheet.Cells(iStartRow, 34).Value = drItem("DO_EAN_CODE")
                    worksheet.Cells(iStartRow, 35).Value = GetValueOrEnumLabel("C_FAM_PRD_WMS", drItem("DO_FAM_PRD_WMS"))
                    worksheet.Cells(iStartRow, 36).Value = GetValueOrEnumLabel("C_FORMATO_PACK", drItem("DO_FORMATO_PACK"))
                    worksheet.Cells(iStartRow, 37).Value = GetValueOrEnumLabel("C_GEST_LOTTO_CLIENTE", drItem("DO_GEST_LOTTO_CLI"))
                    worksheet.Cells(iStartRow, 38).Value = GetValueOrEnumLabel("C_GEST_LOTTO_WMS", drItem("DO_GEST_LOTTO_WMS"))
                    worksheet.Cells(iStartRow, 39).Value = drItem("DO_GG_MEDI_CQ")
                    worksheet.Cells(iStartRow, 40).Value = GetValueOrEnumLabel("C_GRP_CONTAB", drItem("DO_GRP_CONTAB"))
                    worksheet.Cells(iStartRow, 41).Value = GetValueOrEnumLabel("C_GRP_PRD", drItem("DO_GRP_PRD"))
                    worksheet.Cells(iStartRow, 42).Value = GetValueOrEnumLabel("C_IMBALLO_SML", drItem("DO_IMBALLO_SML"))
                    worksheet.Cells(iStartRow, 43).Value = GetValueOrEnumLabel("C_INGREDIENTI", drItem("DO_INGREDIENTI"))
                    worksheet.Cells(iStartRow, 44).Value = GetValueOrEnumLabel("C_LAVORAZIONE", drItem("DO_LAVORAZIONE"))
                    worksheet.Cells(iStartRow, 45).Value = GetValueOrEnumLabel("C_LINEA_BUSINESS", drItem("DO_LINEA_BUSINESS"))
                    worksheet.Cells(iStartRow, 46).Value = GetValueOrEnumLabel("C_LINGUA", drItem("DO_LINGUA"))
                    worksheet.Cells(iStartRow, 47).Value = GetValueOrEnumLabel("C_LINGUA", drItem("DO_LINGUA_ETICHETTA"))
                    worksheet.Cells(iStartRow, 48).Value = GetValueOrEnumLabel("C_MACRO_CERT", drItem("DO_MACRO_CERT"))
                    worksheet.Cells(iStartRow, 49).Value = GetValueOrEnumLabel("C_MARCHIO", drItem("DO_MARCHIO"))
                    
                    if drItem("DO_MERCE_SELEZIONATA").ToString() = "1" then
                        worksheet.Cells(iStartRow, 50).Value = "SELEZIONATA"
                    Elseif drItem("DO_MERCE_SELEZIONATA").ToString() = "0" then 
                        worksheet.Cells(iStartRow, 50).Value = "NON SELEZIONATA"
                    Else
                        worksheet.Cells(iStartRow, 50).Value = "NON COMPILATO"
                    end if
                    
                    worksheet.Cells(iStartRow, 51).Value = GetValueOrEnumLabel("C_MONO_MULTIORIGINE", drItem("DO_MONO_MULTIORIGINE"))
                    worksheet.Cells(iStartRow, 52).Value = GetValueOrEnumLabel("C_ORIGINE", drItem("DO_ORIG_DOGANALE"))
                    worksheet.Cells(iStartRow, 53).Value = GetValueOrEnumLabel("C_ORIGINE", drItem("DO_PAESE_ORIGINE"))
                    worksheet.Cells(iStartRow, 54).Value = GetValueOrEnumLabel("C_RADICE", drItem("DO_RADICE"))
                    worksheet.Cells(iStartRow, 55).Value = GetValueOrEnumLabel("C_REG_ORIGINE", drItem("DO_REGIONE_ORIGINE"))
                    worksheet.Cells(iStartRow, 56).Value = GetValueOrEnumLabel("C_RICETTA", drItem("DO_RICETTA"))
                    worksheet.Cells(iStartRow, 57).Value = GetValueOrEnumLabel("C_SETT_MERCE", drItem("DO_SETT_MERCE"))
                    worksheet.Cells(iStartRow, 58).Value = GetValueOrEnumLabel("C_SINGLE_MULTIBLEND", drItem("DO_SINGLE_MULTIBLEND"))
                    worksheet.Cells(iStartRow, 59).Value = GetValueOrEnumLabel("C_STATUS", drItem("DO_STATUS"))
                    worksheet.Cells(iStartRow, 60).Value = GetValueOrEnumLabel("C_TIPO", drItem("DO_TIPO"))
                    worksheet.Cells(iStartRow, 61).Value = GetValueOrEnumLabel("C_TIPO_IMBALLO", drItem("DO_TIPO_IMBALLO"))
                    worksheet.Cells(iStartRow, 62).Value = GetValueOrEnumLabel("C_TIPO_PRD", drItem("DO_TIPO_PRD"))
                    worksheet.Cells(iStartRow, 63).Value = GetValueOrEnumLabel("C_TIPO_SML", drItem("DO_TIPO_SML"))
                    worksheet.Cells(iStartRow, 64).Value = drItem("DO_VARIANTE")

                    ' CERTIFICATION CATEGORY
                    if drItem("DO_FAIR_TRADE").ToString() = "1" then
                        worksheet.Cells(iStartRow, 65).Value = "FAIR_TRADE"
                    elseif drItem("DO_FAIR_TRADE").ToString() = "0" then
                        worksheet.Cells(iStartRow, 65).Value = "NON FAIR_TRADE"
                    else 
                        worksheet.Cells(iStartRow, 65).Value = "NON COMPILATO"
                    end if
                    
                    if drItem("DO_FFL").ToString() = "1" then
                        worksheet.Cells(iStartRow, 66).Value = "FFL"
                    elseif drItem("DO_FFL").ToString() = "0" then
                        worksheet.Cells(iStartRow, 66).Value = "NON FFL"
                    else 
                        worksheet.Cells(iStartRow, 66).Value = "NON COMPILATO"
                    end if
                    
                    if drItem("DO_FILIERA_CONV").ToString() = "1" then
                        worksheet.Cells(iStartRow, 67).Value = "FILIERA_CONV"
                    Else
                        worksheet.Cells(iStartRow, 67).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FILIERA_CONV")) 
                    end if
                    
                    if drItem("DO_FILIERA_CONV_CERT").ToString() = "1" then
                        worksheet.Cells(iStartRow, 68).Value = "FILIERA_CONV_CERT"
                    Else
                        worksheet.Cells(iStartRow, 68).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FILIERA_CONV_CERT")) 
                    end if
                    
                    if drItem("DO_HALAL").ToString() = "1" then
                        worksheet.Cells(iStartRow, 69).Value = "HALAL"
                    Else
                        worksheet.Cells(iStartRow, 69).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_HALAL")) 
                    end if
                    
                    if drItem("DO_IBD").ToString() = "1" then
                        worksheet.Cells(iStartRow, 70).Value = "IBD"
                    Else
                        worksheet.Cells(iStartRow, 70).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_IBD")) 
                    end if
                    
                    if drItem("DO_JAS").ToString() = "1" then
                        worksheet.Cells(iStartRow, 71).Value = "JAS"
                    Else
                        worksheet.Cells(iStartRow, 71).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_JAS")) 
                    end if
                    
                    if drItem("DO_KOSHER").ToString() = "1" then
                        worksheet.Cells(iStartRow, 72).Value = "KOSHER"
                    Else
                        worksheet.Cells(iStartRow, 72).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_KOSHER")) 
                    end if
                    
                    if drItem("DO_NATURLAND").ToString() = "1" then
                        worksheet.Cells(iStartRow, 73).Value = "NATURLAND"
                    Else
                        worksheet.Cells(iStartRow, 73).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_NATURLAND")) 
                    end if
                    
                    if drItem("DO_NOP").ToString() = "1" then
                        worksheet.Cells(iStartRow, 74).Value = "NOP"
                    Else
                        worksheet.Cells(iStartRow, 74).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_NOP")) 
                    end if
                    
                    if drItem("DO_OIA").ToString() = "1" then
                        worksheet.Cells(iStartRow, 75).Value = "OIA"
                    Else
                        worksheet.Cells(iStartRow, 75).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_OIA")) 
                    end if
                    
                    if drItem("DO_PVB").ToString() = "1" then
                        worksheet.Cells(iStartRow, 76).Value = "PVB"
                    Else
                        worksheet.Cells(iStartRow, 76).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_PVB")) 
                    end if
                    
                    if drItem("DO_REGEN_ORGANIC").ToString() = "1" then
                        worksheet.Cells(iStartRow, 77).Value = "REGEN_ORGANIC"
                    Elseif drItem("DO_REGEN_ORGANIC").ToString() = "0" then 
                        worksheet.Cells(iStartRow, 77).Value = "NON REGEN_ORGANIC"
                    Else
                        worksheet.Cells(iStartRow, 77).Value = "NON COMPILATO"
                    end if
                    
                    if drItem("DO_SPIGA_BARRATA").ToString() = "1" then
                        worksheet.Cells(iStartRow, 78).Value = "SPIGA_BARRATA"
                    Else
                        worksheet.Cells(iStartRow, 78).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_SPIGA_BARRATA")) 
                    end if
                    
                    if drItem("DO_VEGAN").ToString() = "1" then
                        worksheet.Cells(iStartRow, 79).Value = "VEGAN"
                    Else
                        worksheet.Cells(iStartRow, 79).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_VEGAN")) 
                    end if
                    
                    if drItem("DO_BIO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 80).Value = "BIO"
                    Else
                        worksheet.Cells(iStartRow, 80).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_BIO")) 
                    end if
                    
                    if drItem("DO_BIO_SUISSE").ToString() = "1" then
                        worksheet.Cells(iStartRow, 81).Value = "BIO SUISSE"
                    Else
                        worksheet.Cells(iStartRow, 81).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_BIO_SUISSE")) 
                    end if
                    
                    if drItem("DO_CONVENZIONALE").ToString() = "1" then
                        worksheet.Cells(iStartRow, 82).Value = "CONVENZIONALE"
                    Else
                        worksheet.Cells(iStartRow, 82).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_CONVENZIONALE")) 
                    end if
                    
                    if drItem("DO_CMO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 83).Value = "CMO"
                    Else
                        worksheet.Cells(iStartRow, 83).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_CMO")) 
                    end if
                    
                    if drItem("DO_DEMETER").ToString() = "1" then
                        worksheet.Cells(iStartRow, 84).Value = "DEMETER"
                    Else
                        worksheet.Cells(iStartRow, 84).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_DEMETER")) 
                    end if
                    
                    'QUALITA CATAGORY
                    if drItem("DO_ALLERGENI").ToString() = "1" then
                        worksheet.Cells(iStartRow, 85).Value = "ALLERGENI"
                    Else
                        worksheet.Cells(iStartRow, 85).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_ALLERGENI")) 
                    end if
                   
                    if drItem("DO_BIOLOGICO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 86).Value = "BIOLOGICO"
                    Else
                        worksheet.Cells(iStartRow, 86).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_BIOLOGICO")) 
                    end if
                    
                    if drItem("DO_CAR_COLORE_A").ToString() = "1" then
                        worksheet.Cells(iStartRow, 87).Value = "DEMETER"
                    Else
                        worksheet.Cells(iStartRow, 87).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_CAR_COLORE_A")) 
                    end if
                    
                    worksheet.Cells(iStartRow, 88).Value = GetValueOrEnumLabel("C_DO_YESNO", drItem("DO_CAR_COLORE_B"))
                    worksheet.Cells(iStartRow, 89).Value = GetValueOrEnumLabel("C_DO_YESNO", drItem("DO_CAR_COLORE_L"))
                    worksheet.Cells(iStartRow, 90).Value = GetValueOrEnumLabel("C_CONAI", drItem("DO_CODICE_CONAI"))
                    worksheet.Cells(iStartRow, 91).Value = drItem("DO_DEN_CONV_UDB")
                    
                    if drItem("DO_FREDDO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 92).Value = "FREDDO"
                    Else
                        worksheet.Cells(iStartRow, 92).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FREDDO")) 
                    end if
                    
                    if drItem("DO_LOTTO_PROPR_CLI").ToString() = "1" then
                        worksheet.Cells(iStartRow, 93).Value = "LOTTO PROPR CLI"
                    Else
                        worksheet.Cells(iStartRow, 93).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_LOTTO_PROPR_CLI")) 
                    end if
                    
                    worksheet.Cells(iStartRow, 94).Value = GetValueOrEnumLabel("C_MATERIALE_IMBALLO", drItem("DO_MATERIALE_IMBALLO"))
                    worksheet.Cells(iStartRow, 95).Value = drItem("DO_NUM_CONV_UDB")
                    worksheet.Cells(iStartRow, 96).Value = drItem("DO_PESO_LORDO")
                    worksheet.Cells(iStartRow, 97).Value = drItem("DO_PESO_NETTO")
                    
                    if drItem("DO_SECONDO_CONTROLLO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 98).Value = "SECONDO CONTROLLO"
                    Else
                        worksheet.Cells(iStartRow, 98).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_SECONDO_CONTROLLO")) 
                     end if
                
                    worksheet.Cells(iStartRow, 99).Value = drItem("DO_SIZE")
                    
                    if drItem("DO_SLUNGA_COMPLIANT").ToString() = "1" then
                        worksheet.Cells(iStartRow, 100).Value = "ESSELUNGA COMPLIANT"
                    Else
                        worksheet.Cells(iStartRow, 100).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_SLUNGA_COMPLIANT")) 
                    end if
                    
                    if drItem("DO_SMT").ToString() = "1" then
                        worksheet.Cells(iStartRow, 101).Value = "SMT"
                    Else
                        worksheet.Cells(iStartRow, 101).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_SMT")) 
                    end if
                    
                    worksheet.Cells(iStartRow, 102).Value = drItem("DO_TEMP_TOSTATURA")
                    
                    if drItem("DO_TRACCIABILE").ToString() = "1" then
                        worksheet.Cells(iStartRow, 103).Value = "TRACCIABILE"
                    Else
                        worksheet.Cells(iStartRow, 103).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_TRACCIABILE")) 
                    end if
                    
                    worksheet.Cells(iStartRow, 104).Value = GetValueOrEnumLabel("C_UDM_ALT_ISO", drItem("DO_UDM_ALT_ISO"))
                    worksheet.Cells(iStartRow, 105).Value = drItem("SHELF_LIFE")
    
                    'REPORT BIO CATAGORY
                    worksheet.Cells(iStartRow, 106).Value = GetValueOrEnumLabel("C_FAM_REPORT_BIO_SPP", drItem("DO_FAM_REPORT_BIO_SP"))
                    
                    if drItem("DO_IN_REPORT_BIO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 107).Value = "IN REPORT BIO"
                    Else
                        worksheet.Cells(iStartRow, 107).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_IN_REPORT_BIO")) 
                    end if
                    
                    if drItem("DO_IN_REPORT_BIO_SPP").ToString() = "1" then
                        worksheet.Cells(iStartRow, 108).Value = "IN REPORT BIO SPP"
                    Else
                        worksheet.Cells(iStartRow, 108).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_IN_REPORT_BIO_SPP")) 
                    end if
                    
                    'BLANK(NO) CATAGORY
                    if drItem("DO_CONV_CERTIFICATO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 109).Value = "CONV_CERTIFICATO"
                    Else
                        worksheet.Cells(iStartRow, 109).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_CONV_CERTIFICATO")) 
                    end if
                    
                    if drItem("DO_DAMIANO_SUPER_DOP").ToString() = "1" then
                        worksheet.Cells(iStartRow, 110).Value = "DAMIANO_SUPER_DOP"
                    Else
                        worksheet.Cells(iStartRow, 110).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_DAMIANO_SUPER_DOP")) 
                    end if
                    
                    if drItem("DO_DOP").ToString() = "1" then
                        worksheet.Cells(iStartRow, 111).Value = "DOP"
                    Else
                        worksheet.Cells(iStartRow, 111).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_DOP")) 
                    end if
                    
                    if drItem("DO_DOP_BIO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 112).Value = "DOP_BIO"
                    Else
                        worksheet.Cells(iStartRow, 112).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_DOP_BIO")) 
                    end if
                    
                    if drItem("DO_FILIERA_BIO_GOLD").ToString() = "1" then
                        worksheet.Cells(iStartRow, 113).Value = "FILIERA_BIO_GOLD"
                    Else
                        worksheet.Cells(iStartRow, 113).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FILIERA_BIO_GOLD")) 
                    end if
                    
                    if drItem("DO_FILIERA_BIO_PREM").ToString() = "1" then
                        worksheet.Cells(iStartRow, 114).Value = "FILIERA_BIO_PREM"
                    Else
                        worksheet.Cells(iStartRow, 114).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FILIERA_BIO_PREM")) 
                    end if
                    
                    if drItem("DO_FILIERA_BIO_STAND").ToString() = "1" then
                        worksheet.Cells(iStartRow, 115).Value = "FILIERA_BIO_STAND"
                    Else
                        worksheet.Cells(iStartRow, 115).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_FILIERA_BIO_STAND")) 
                    end if
                    
                    if drItem("DO_IGP").ToString() = "1" then
                        worksheet.Cells(iStartRow, 116).Value = "IGP"
                    Else
                        worksheet.Cells(iStartRow, 116).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_IGP")) 
                    end if
                    
                    if drItem("DO_IGP_BIO").ToString() = "1" then
                        worksheet.Cells(iStartRow, 117).Value = "IGP_BIO"
                    Else
                        worksheet.Cells(iStartRow, 117).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_IGP_BIO")) 
                    end if
                    
                    if drItem("DO_NESSUNA_CERT_EST").ToString() = "1" then
                        worksheet.Cells(iStartRow, 118).Value = "NESSUNA_CERT_EST"
                    Else
                        worksheet.Cells(iStartRow, 118).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_NESSUNA_CERT_EST")) 
                    end if
                    
                    if drItem("DO_NESSUNA_CERT_INT").ToString() = "1" then
                        worksheet.Cells(iStartRow, 119).Value = "NESSUNA_CERT_INT"
                    Else
                        worksheet.Cells(iStartRow, 119).Value = GetValueOrEnumLabel("C_DO_YESNO",drItem("DO_NESSUNA_CERT_INT")) 
                    end if
                    
                    'setting the borders
					worksheet.Cells(iStartRow , 1 , iStartRow , 119).Style.Border.Left.Style = ExcelBorderStyle.Thin
					worksheet.Cells(iStartRow , 1 , iStartRow , 119).Style.Border.Top.Style = ExcelBorderStyle.Thin
					worksheet.Cells(iStartRow , 1 , iStartRow , 119).Style.Border.Right.Style = ExcelBorderStyle.Thin
					worksheet.Cells(iStartRow , 1 , iStartRow , 119).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
				
				worksheet.Row(iStartRow).Height = 30
				iStartRow = iStartRow + 1
				next
			next
		End if
		
		'printing the report with idm
		Dim filename as string = "Damianos Item" & _WIPID & ".xlsx"
		Dim oidm as IDM_EPPLUS = New IDM_EPPLUS(Me)
		filename = filename.Replace(".xlsx" , "")
	
		Catch ex as exception
			messagelist("Report Script: " + ex.Message)
			Return 9111
		End Try
		Return 111
	End Function

Function getEnumLabel(byval enumList as string, byval enumValue as string) as string
    Dim dsList As DataSet = ObjectDataSet("ENUMLIST", enumList, "CF")
	Dim EnumTableName As String = DataSetTableName("ENUMLIST", enumList, "CF")
	Dim drVal() As Data.DataRow = dsList.Tables(EnumTableName).Select("ENUM_VALUE = '"+ enumValue +"'")
	'messagelist(enumValue)
	Dim enumlabel As String = ""
	 If drVal IsNot Nothing AndAlso drVal.Length > 0 Then
        Dim enumLabelObject As Object = drVal(0)("ENUM_LABEL2")
        If enumLabelObject IsNot DBNull.Value Then
            enumLabel = enumLabelObject.ToString()
            'messagelist(enumLabel)
        End If
    End if
    return enumlabel
End Function

Function GetValueOrEnumLabel(ByVal enumList As String, ByVal enumValue As Object) As String
    If enumValue Is DBNull.Value OrElse String.IsNullOrEmpty(enumValue.ToString()) Then
        ' Special case for "C_DO_YESNO" when enumValue is empty
        If enumList = "C_DO_YESNO" Then
            Return "NON COMPILATO"
        end If
        
    Else
        
        If enumList = "C_DO_YESNO" AndAlso enumValue.ToString() = "0" Then
            Return ""
        else 
            Return getEnumLabel(enumList, enumValue.ToString()) 
        end If
        
    End If
End Function

End Class       
	