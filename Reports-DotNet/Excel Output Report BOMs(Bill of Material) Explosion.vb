' Purpose:  Print Formula's Nutrituin Panel Facts

' AUTHOR    :   Reza     January 2025

' NOTE: 	This file includes only the custom-developed logic for creating the Excel Output report for BOMs Explosion
' 			It does not contain full Interface Configuration or system templates generated.

Option Strict Off
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
Imports System.Collections.Specialized
Imports System.Text.RegularExpressions
Imports System.Xml.Linq
Imports System.Data.Common

Class ActionScript
	Inherits FcProcFuncSetEventWF
	
	'for retreiving the status label
	dim gf as GENERALFUNCTIONS = New GENERALFUNCTIONS(Me) 
	
	Function wf_start() As Long
	
	Try
	
		'Create Excel Spreadsheet File
		Dim pkg as New ExcelPackage()
		
		'Add a worksheet to the Excel Spreadsheet
		Dim worksheet As ExcelWorksheet = pkg.workbook.Worksheets.Add("Formulas Components")
		Dim rowIndex As Integer = 0
		
		'Define the default font and size
		worksheet.Cells.Style.font.Size= 10
		
		Dim ColumnWidth() As Double = {6,20,18,7,25,10,9,15,50,50,9}
		
		For I as Integer = 1 To columnWidth.Length
			Worksheet.Column(i).Width = columnWidth(i - 1)
			Worksheet.Column(i).Style.WrapText = True
			Worksheet.Column(i).Style.VerticalAlignment = ExcelVerticalAlignment.Center
		Next i
		
		' Set the first title
        Worksheet.Cells(1, 1).Value = "FORMULA " & _OBJECTKEY & " : " & objproperty("DESCRIPTION")
        Worksheet.Cells(1, 1).Style.Font.Bold = True
        Worksheet.Cells(1, 1, 1, 11).Merge = True
        Worksheet.Cells(1, 1, 1, 11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
        
        Dim frmStatus as String = objproperty("STATUSIND.STATUS") 
        Worksheet.Cells(2 , 1).Value = "Current Status of Formula: " & gf.GetStatusDesc("FORMULA",frmStatus,"EN-US")
		Worksheet.Cells(2, 1, 2, 11).Merge = True
		Worksheet.Cells(2, 1, 2, 11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
		
		'Formula Location
		Dim sWorkcenters As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "" , "*", 1)
        Dim sWorkcentersDesc As Object = ObjProperty("DESCRIPTION.CONTEXT", "", "" , "*", 1)
        Dim formSites as String = ""
        For y as Integer = 0 To sWorkcenters.Length -1 
            formSites =  sWorkcenters(y) + "-" + sWorkcentersDesc(y) + " ; " + formSites
        Next
        If formSites.EndsWith(" ; ") Then
            formSites = formSites.Substring(0, formSites.Length - 3)
        End If
        
        Dim frmAllergens as string = ""
		Dim parameters() as String = {"ALLERGEN_WHEAT","ALLERGEN_CRUSTACEA","ALLERGEN_EGG","ALLERGEN_FISH","ALLERGEN_PEANUTS","ALLERGEN_SOY","ALLERGEN_MILK","ALLERGEN_TREENUTS","ALLERGEN_CELERY","ALLERGEN_MUSTARD","ALLERGEN_SESAME","ALLERGEN_SO2","ALLERGEN_LUPIN","ALLERGEN_MOLLUSCS","28_SULPHITES_NUM","GLUTEN_CONTENT"}
        For Each param As String In parameters
            Dim paramvalue As Object = ObjProperty("VALUE.TPALL", "FORMULA", "" , param, "PARAM_CODE")
            If IsBlank(paramvalue) = 0 Then
                If param.StartsWith("ALLERGEN_") Then
                    If paramvalue = "4" Then
                        frmAllergens = param + " ; " + frmAllergens
                    End if
                Else 
                    Dim checkValue as Double = Math.Round(CDbl(paramvalue), 1 , MidpointRounding.AwayFromZero)
                    If checkValue <> 0 Then
                        frmAllergens = ( param + " = " + cstr(checkValue) + "ppm") + " ; " +frmAllergens 
                    End if 
                End if 
            End if 
        Next 
        
        If frmAllergens.EndsWith(" ; ") Then
            frmAllergens = frmAllergens.Substring(0, frmAllergens.Length - 3)
        End If
        
        
		Worksheet.Cells(3 , 1).Value = "Manufacturing Location: " + formSites
		Worksheet.Cells(3, 1, 3, 11).Merge = True
		Worksheet.Cells(3, 1, 3, 11).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
		
		Worksheet.Cells(4, 1, 4, 11).Merge = True
	    
		'Headers' Title
		worksheet.Cells("A5").Value = "Level"
		worksheet.Cells("B5").Value = "FG/WIP"
        worksheet.Cells("C5").Value = "Item code"
        worksheet.Cells("D5").Value = "Item Uom"
        worksheet.Cells("E5").Value = "Description"
        worksheet.Cells("F5").Value = "Class"       
        worksheet.Cells("G5").Value = "Current Status"            
        worksheet.Cells("H5").Value = "Status Desc"
        worksheet.Cells("I5").Value = "Site(s)"
        worksheet.Cells("J5").Value = "Allergens"
        worksheet.Cells("K5").Value = "Shelf Life (day)"

        'setting of the costumization   
        worksheet.Cells("A1:K5").Style.Font.Size = 11
		worksheet.Cells("A1:K5").Style.Font.Bold = True
		worksheet.Cells("A1:K6").Style.Border.Left.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A1:K6").Style.Border.Top.Style = ExcelBorderStyle.Thin
    	worksheet.Cells("A1:K6").Style.Border.Right.Style = ExcelBorderStyle.Thin
		worksheet.Cells("A1:K6").Style.Border.Bottom.Style = ExcelBorderStyle.Thin
		
		'Starting for the inputs 
		Dim sXML as string
    	sXML = ObjectXML("",  "", "HEADER;CONTEXT;LINELIST")
    	
    	Dim xDoc as XDocument = XDocument.Parse(sXML)
    	
'*******Adding Parent_Code Node to XML with the value of the parentformula of each itemIngred
    	Dim allItems As IEnumerable(Of XElement) =
          From el In xDoc...<fsxml>.<FSFORMULA>.<FSFORMULALINELIST>
          Select el
              
        'sXML = xDoc.ToString()
		'messagelist(sXML)
		
        Dim currentParentPerLevel As New Dictionary(Of Integer, String)
        Dim parentNode  As XElement
        
        For Each ingred As XElement In allItems
            Dim levelStr as string = ingred.<LEVEL>.Value
            Dim itemCode as string = ingred.<ITEM_CODE>.Value
            Dim formulaCodeNode As string = ingred.<FORMULA_CODE>.Value
            Dim classCode as string = ingred.<CLASS>.Value
            Dim level As Integer = CInt(levelStr)
            
            Dim parentCode As String = ""
            
            If currentParentPerLevel.ContainsKey(level - 1) Then
              parentCode = currentParentPerLevel(level - 1)
            ElseIf currentParentPerLevel.ContainsKey(level) Then
              parentCode = currentParentPerLevel(level)
            Else
              parentCode = _OBJECTKEY
            End If
            
            if level = 0 then parentCode = _OBJECTKEY
            
            ingred.Add(New XElement("PARENT_CODE", parentCode))
               
           If formulaCodeNode IsNot Nothing Then
               currentParentPerLevel(level) = formulaCodeNode
           End If
        Next
    	
'*******Sorting the XML according to Level and ComponentInd
        Dim ItemsList As IEnumerable(Of XElement) =
            From el In xDoc...<fsxml>.<FSFORMULA>.<FSFORMULALINELIST>
            Order By CInt(el.<LEVEL>.Value) Ascending, CInt(el.<COMPONENT_IND>.Value) Ascending
            Select el

        Dim itemCodes as string = ""
        Dim itemUom as String = ""
        Dim formulaCode as String = ""
        Dim primeLevel as String = ""
        Dim nextLevel as String = ""
        Dim itemDesc as string = ""
        Dim itemLevel as string = ""
        Dim itemStatus as string = ""
        Dim itemStatDesc as string = ""
        Dim itemClass as string = ""
        Dim iStartRow As Integer = 7
        Dim checkYellow as Integer = 0
        Dim parentFormula As String = ""
        Dim hierarchyId as String = ""
        Dim levelFormulaMap As New Dictionary(Of Integer, String)
        
'*******set values of the current Formula as the fisrs row
        worksheet.Cells("A6").Value = "0"
		worksheet.Cells("B6").Value = _OBJECTKEY
        worksheet.Cells("C6").Value = ObjProperty("ITEMCODE")
        worksheet.Cells("D6").Value = ObjProperty("UOMCODE","ITEM",ObjProperty("ITEMCODE"))
        worksheet.Cells("E6").Value = ObjProperty("DESCRIPTION")
        worksheet.Cells("F6").Value = ObjProperty("CLASS")
        worksheet.Cells("G6").Value = frmStatus           
        worksheet.Cells("H6").Value = gf.GetStatusDesc("FORMULA", frmStatus ,"EN-US")
        worksheet.Cells("I6").Value = formSites
        worksheet.Cells("J6").Value = frmAllergens
        worksheet.Cells("K6").Value = ObjProperty("VALUE.TPALL", "FORMULA", "" , "SHELF_LIFE", "PARAM_CODE")
        
'*******Loop items list to fill the other rows       
		For Each itemLine As XElement In ItemsList
		    
		    hierarchyId = itemLine.<ID_HIERARCHY>.Value
		    parentFormula = itemLine.<PARENT_CODE>.Value
		    If not ((hierarchyId.Split("."c).Length - 1 > 1) AndAlso (parentFormula = _OBJECTKEY)) Then
		    
    		    formulaCode = itemLine.<FORMULA_CODE>.Value
    			itemCodes = itemLine.<ITEM_CODE>.Value
    			itemUom = ""
    			itemDesc = itemLine.<DESCRIPTION>.Value
    			itemLevel = CInt(itemLine.<LEVEL>.Value) 
    			itemClass = itemLine.<CLASS>.Value
    			itemStatus = itemLine.<STATUS>.Value
    			parentFormula = itemLine.<PARENT_CODE>.Value
    			Dim shelfLife as String = ""
    		
    			
    '***********Check the sites and allergens of the BOMS for filling the columns Sites and Allergens and ItemUom
    			Dim itemAllergens as string = ""
    			Dim itemSite As String = ""
    			Dim params() as String = {"ALLERGEN_WHEAT","ALLERGEN_CRUSTACEA","ALLERGEN_EGG","ALLERGEN_FISH","ALLERGEN_PEANUTS","ALLERGEN_SOY","ALLERGEN_MILK","ALLERGEN_TREENUTS","ALLERGEN_CELERY","ALLERGEN_MUSTARD","ALLERGEN_SESAME","ALLERGEN_SO2","ALLERGEN_LUPIN","ALLERGEN_MOLLUSCS","28_SULPHITES_NUM","GLUTEN_CONTENT"}
    			
    			If Not String.IsNullOrEmpty(formulaCode) AndAlso formulaCode <> "0" Then
    			    Dim sfrmWorkCenters as Object = ObjProperty("ATTRIBVAL.CONTEXT", "FORMULA", formulaCode , "*", 1)
    			    Dim sfrmWorkCentersDesc as Object = ObjProperty("DESCRIPTION.CONTEXT", "FORMULA", formulaCode , "*", 1)
    			    For y As Integer = 0 To sfrmWorkCenters.Length - 1
                        itemSite = sfrmWorkCenters(y) & "-" & sfrmWorkCentersDesc(y) & " ; " & itemSite       
    			    Next
                    If itemSite.EndsWith(" ; ") Then
                        itemSite = itemSite.Substring(0, itemSite.Length - 3)
                    End If
                    For Each param As String In params
                        Dim paramvalue As Object = ObjProperty("VALUE.TPALL", "FORMULA", formulaCode , param, "PARAM_CODE")
                        If IsBlank(paramvalue) = 0 Then
                            If param.StartsWith("ALLERGEN_") Then
                                If paramvalue = "4" Then
                                    itemAllergens = param + " ; " + itemAllergens
                                End if
                            Else 
                                Dim checkValue as Double = Math.Round(CDbl(paramvalue), 1 , MidpointRounding.AwayFromZero)
                                If checkValue <> 0 Then
                                    itemAllergens = ( param + " = " + cstr(checkValue) + "ppm") + " ; " + itemAllergens 
                                End if 
                            End if 
                        End if 
                    Next 
                    
                    If itemAllergens.EndsWith(" ; ") Then
                        itemAllergens = itemAllergens.Substring(0, itemAllergens.Length - 3)
                    End If
                    
    			    itemStatDesc = gf.GetStatusDesc("FORMULA",itemStatus,"EN-US")
    			    
                    itemUom = ObjProperty("UOMCODE","ITEM",itemCodes)
                    
    		    Else
    		        
                    Dim sWarehouses As Object = ObjProperty("ATTRIBVAL.CONTEXT", "ITEM", itemCodes , "*", 1)
                    Dim sWarehousesDesc As Object = ObjProperty("DESCRIPTION.CONTEXT", "ITEM", itemCodes , "*", 1)
                    
                    For y As Integer = 0 To sWarehouses.Length - 1
                        itemSite = sWarehouses(y) & "-" & sWarehousesDesc(y) & " ; " & itemSite       
                    Next
                    If itemSite.EndsWith(" ; ") Then
                        itemSite = itemSite.Substring(0, itemSite.Length - 3)
                    End If
        
                    'Check the Allergens of the itemIngred   
                    For Each param As String In params
                        Dim paramvalue As Object = ObjProperty("VALUE.TPALL", "ITEM", itemCodes , param, "PARAM_CODE")
                        If IsBlank(paramvalue) = 0 Then
                            
                            If param.StartsWith("ALLERGEN_") Then
                                If paramvalue = "4" Then
                                    itemAllergens = param + " ; " + itemAllergens
                                End if
                            Else 
                                Dim checkValue as Double = Math.Round(CDbl(paramvalue), 1 , MidpointRounding.AwayFromZero)
                                If checkValue <> 0 Then
                                    itemAllergens = ( param + " = " + cstr(checkValue) + "ppm") + " ; " + itemAllergens 
                                    
                                End if 
                            End if 
                        End if 
                    Next 
                    
                    If itemAllergens.EndsWith(" ; ") Then
                        itemAllergens = itemAllergens.Substring(0, itemAllergens.Length - 3)
                    End If
                    
                    itemStatDesc = gf.GetStatusDesc("ITEM",itemStatus,"EN-US")
                    
                    itemUom = ObjProperty("UOMCODE","ITEM",itemCodes)
                    
    			end If
    			
    '***********Check the Shelf Life for FG formulas
    			if itemClass = "ITEM_PACK" Then 
    			    shelfLife = ObjProperty("VALUE.TPALL", "FORMULA", formulaCode , "SHELF_LIFE" , "PARAM_CODE")
    			end if 
    		    
    			worksheet.Cells(iStartRow, 1).Value = itemLevel
                worksheet.Cells(iStartRow, 2).Value = parentFormula
                worksheet.Cells(iStartRow, 3).Value = itemCodes
                worksheet.Cells(iStartRow, 4).Value = itemUom
                worksheet.Cells(iStartRow, 5).Value = itemDesc
                worksheet.Cells(iStartRow, 6).Value = itemClass
                worksheet.Cells(iStartRow, 7).Value = itemStatus
                worksheet.Cells(iStartRow, 8).Value = itemStatDesc
                worksheet.Cells(iStartRow, 9).Value = itemSite
                worksheet.Cells(iStartRow, 10).Value = itemAllergens
                worksheet.Cells(iStartRow, 11).Value = shelfLife
                
    '***********Apply alternating row colors (blue and white)
                If iStartRow Mod 2 = 0 Then
                    worksheet.Cells(iStartRow, 1, iStartRow, 11).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(iStartRow, 1, iStartRow, 11).Style.Fill.BackgroundColor.SetColor(Color.LightBlue) ' Blue rows
                Else
                    worksheet.Cells(iStartRow, 1, iStartRow, 11).Style.Fill.PatternType = ExcelFillStyle.Solid
                    worksheet.Cells(iStartRow, 1, iStartRow, 11).Style.Fill.BackgroundColor.SetColor(Color.White) ' White rows
                End If
                
    '***********Check the row for the ones which are mfgItems to be yellow
                if iStartRow > 6 then 
                    If (itemClass <> "RM" AndAlso itemClass <> "PKG") Then
        			    worksheet.Cells(iStartRow , 1, iStartRow , 11).Style.Fill.PatternType = ExcelFillStyle.Solid
                        worksheet.Cells(iStartRow , 1, iStartRow , 11).Style.Fill.BackgroundColor.SetColor(Color.LightYellow)
                    END if
    			end if
                
                'setting the borders
    			worksheet.Cells(iStartRow , 1 , iStartRow , 11).Style.Border.Left.Style = ExcelBorderStyle.Thin
    			worksheet.Cells(iStartRow , 1 , iStartRow , 11).Style.Border.Top.Style = ExcelBorderStyle.Thin
    			worksheet.Cells(iStartRow , 1 , iStartRow , 11).Style.Border.Right.Style = ExcelBorderStyle.Thin
    			worksheet.Cells(iStartRow , 1 , iStartRow , 11).Style.Border.Bottom.Style = ExcelBorderStyle.Thin
    			worksheet.Row(iStartRow).Height = 30
    			
    			'set the Allignments
    			worksheet.Cells("A6:A" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right       ' Level
                worksheet.Cells("B6:B" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Formula
                worksheet.Cells("C6:C" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Item Code
                worksheet.Cells("D6:D" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Item Uom
                worksheet.Cells("E6:E" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Description
                worksheet.Cells("F6:F" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Class
                worksheet.Cells("G6:G" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right       ' Status
                worksheet.Cells("H6:H" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Status Desc
                worksheet.Cells("I6:I" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Sites(s)
                worksheet.Cells("J6:J" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left        ' Allergens
                worksheet.Cells("K6:K" & iStartRow).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right       ' Shelf Life
                
    		    iStartRow = iStartRow + 1
    		    
    		    'messagelist("formulaCode: " + formulaCode + "ItemCode: " + itemCodes + " , Item DESC:" + itemDesc + " , Item Class: " + itemClass + " , Item Level: " + itemLevel + " , ITEM Status:" + itemStatus + " , Status Desc: " + itemStatDesc )
    		    
		    else 
		        itemCodes = itemLine.<ITEM_CODE>.Value
		        messagelist(itemCodes)
    		    
		    end if
		    
		Next
		
		'Set always the row 6 to LightYellow
        worksheet.Cells(6, 1, 6, 11).Style.Fill.PatternType = ExcelFillStyle.Solid
        worksheet.Cells(6, 1, 6, 11).Style.Fill.BackgroundColor.SetColor(Color.LightYellow)
		
		sXML = xDoc.ToString()
		'messagelist(sXML)
		
		Catch ex as exception
			messagelist("Report Script: " + ex.Message)
			Return 9111
		End Try
		Return 111
	End Function

End Class       
	