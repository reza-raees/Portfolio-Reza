'Description: Massive workflow for automatic calculation of parameters of all formulas 
'Note: first the first workflow is launching and then for each formula in the defined table the second workflow will be launched

'========================================================================================================================================================
'Start : Reza - first workflow - This is the automated Workflow for massive calculation of the Allergen Parameters of all formulas in the table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================
Option Strict Off
imports System
imports System.Diagnostics
imports Microsoft.Visualbasic
imports System.Data
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Threading
Imports System.Linq
Imports Formation.Shared.Defs
Imports FsSrvCore
Imports System.IO
Imports System.Text

Class ActionScript
 Inherits FcProcFuncSetEventWF
 
Function wf_start() As Long

Return 1

End Function

Function wf_complete() As Long
    
    try
        Dim dt as Datatable = TableLookupEx("SO_MASSIVE_ALLERGENS_CALC","TableLookup")
        
        for each row as DataRow in dt.Rows
            Dim formula_code as String = row("FORMULA_CODE")
			Dim formula_version as String = row("VERSION")
			Dim formula_key as String = formula_code + "\" + formula_version
            StartWorkflow("SO_CALC_ALLERGENS_FORMULA", "FORMULA", formula_key )
            PublishDebuggingInfo(row("FORMULA_CODE"))
            'messagelist(formula_key)
            PublishDebuggingInfo(row("FORMULA_CODE")+"HAS UPDATED")
        Next
        
        catch ex as exception
        messagelist("Error: ", ex.message)
        PublishDebuggingInfo("The action has failed"+ex.message )
        
    end try


End Function


End Class
'========================================================================================================================================================
'End : Reza - First Workflow - This is the automated Workflow for massive calculation of the Allergen Parameters of all formulas in the table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================

'========================================================================================================================================================
'Begin : Reza - Second Workflow - This is the automated Workflow for massive calculation of the Allergen Parameters of all formulas in the table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================
Option Strict Off
imports System
imports System.Diagnostics
imports Microsoft.Visualbasic
imports System.Data
Imports System.Collections
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Threading
Imports System.Linq
Imports Formation.Shared.Defs
Imports FsSrvCore
Imports System.IO
Imports System.Text

Class ActionScript
 Inherits FcProcFuncSetEventWF

Function wf_start() As Long
        Try
        
        Dim ingritemCodes As Object = objproperty("ITEMCODE.INGR", "", "", "*", 1)
        Dim frmcodeingred as Object = objproperty("FORMULACODE.INGR", "", "", "*", 1)
        Dim itmQty As Object = objproperty("QUANTITY.INGR", "" , "" ,"*",1)
    
        Dim allergenParams As New List(Of String) From {"ALLERG_OTHER_NUTS", "ALLERGEN_CELERY", "ALLERGEN_CEREAL", "ALLERGEN_CRUSTACEA",
            "ALLERGEN_EGG", "ALLERGEN_FISH", "ALLERGEN_LUPIN", "ALLERGEN_MILK", "ALLERGEN_MOLLUSC","ALLERGEN_MUSTARD", "ALLERGEN_NUTS", "ALLERGEN_PEANUTS", "ALLERGEN_SESAME", "ALLERGEN_SOY", "ALLERGEN_SULPHUR_D"}
            
        '===========================================================================
        'Begin: Calculation for Formula Allergens by item's allergens 
        '===========================================================================
        
        Dim allergenValues As New Dictionary(Of String, List(Of Integer)) ' store the values for each allergen
        
        if ingritemCodes.length > 0 then
        
            For x  as integer = 0 to ingritemCodes.length -1
                For Each allergen as string In allergenParams
                    Dim allergenValue As Object = ObjProperty("VALUE.TPALL", "ITEM", ingritemCodes(x), allergen, "PARAM_CODE")
                    
                    ' Convert the allergenValue to Integer
                    Dim value As Integer = If(IsNumeric(allergenValue), CInt(allergenValue), 0)
                    
                    ' Ensure the allergen key exists in the dictionary before adding values
                    If Not allergenValues.ContainsKey(allergen) Then
                        allergenValues(allergen) = New List(Of Integer)()
                    End If
                    
                    ' Add the value to the list for the current allergen
                    allergenValues(allergen).Add(value)
                Next
            Next
            For Each allergen as string In allergenParams
                ' Get the list of values for the current allergen
                Dim values As List(Of Integer) = allergenValues(allergen)
                
                ObjPropertySet( -1 , 0 ,"VALUE.TPALL","","", allergen , "PARAM_CODE")
                If values.Contains(6) AndAlso values.Contains(7) Then
                    ObjPropertySet( 7 , 0 ,"VALUE.TPALL","","", allergen , "PARAM_CODE")
                ElseIf values.Contains(6) Then
                    ObjPropertySet( 6 , 0 ,"VALUE.TPALL","","", allergen , "PARAM_CODE")
                ElseIf values.Contains(7) Then
                    ObjPropertySet( 7 , 0 ,"VALUE.TPALL","","", allergen , "PARAM_CODE")
                End If
            Next
            'MessageList(objproperty("FORMULA_CODE")& "has calculated")
            PublishDebuggingInfo(("FORMULA_CODE")& "has calculated")
        else
            'messagelist(objproperty("FORMULA_CODE")&" does NOT have ingredients")
            PublishDebuggingInfo(objproperty("FORMULA_CODE")+" does NOT have ingredients")
                
        end If
        
        '===========================================================================
        'End: Calculation for Formula Allergens by item's allergens 
        '===========================================================================
        
        '===========================================================================
        'Begin: Calculation for Formulas Allergens ppm MAX  
        '===========================================================================
        Dim allergenppmxvalues As New Dictionary(Of String, Double)
        
        if ingritemCodes.length > 0 then
            
            For x As Integer = 0 To ingritemCodes.Length - 1
            
                ' Retrieve the item class to determine if it is raw material
                Dim itmclass As String = objproperty("COMPONENTIND", "ITEM", ingritemCodes(x))
                
                ' Loop through each allergen
                For Each allergen As String In allergenParams
                    If itmclass = "8" Then
                        
                        Dim allergenValue As Object = ObjProperty("VALUE.TPALL", "ITEM", ingritemCodes(x), allergen, "PARAM_CODE")
                        if (allergenValue = "7" or allergenValue = "6") then
                            ' Retrieve the ppm value for the current allergen for the current ingredient item
                            Dim allergenppmX As Object = ObjProperty("ATTRIBUTE17.TPALL", "ITEM", ingritemCodes(x), allergen, "PARAM_CODE")
                            If IsBlank(allergenppmX) = 0 Then
                                'MessageList(allergen & " ppmX value is: " & allergenppmX)
                    
                                ' Calculate the product of the allergen ppm and the item quantity
                                Dim multipAllergenppmX As Double = Math.Round((CDbl(itmQty(x)) * CDbl(allergenppmX) / 1000), 8, MidpointRounding.AwayFromZero)
                                'MessageList("Multiplication of quantity and " & allergen & " ppmX is: " & multipAllergenppmX)
                    
                                ' Store or accumulate the value in the dictionary for the current allergen
                                If allergenppmxvalues.ContainsKey(allergen) Then
                                    allergenppmxvalues(allergen) += multipAllergenppmX
                                Else
                                    allergenppmxvalues.Add(allergen, multipAllergenppmX)
                                End If
                            'Else 
                                'MessageList("Calculation stopped because Allergen ppmX value for item " & ingritemCodes(x) & " and allergen " & allergen & " is null")
                                'checkpoint = True
                            End If
                        End If
                        
                    ElseIf itmclass = "1" Then
                        
                        Dim allergenValue As Object = ObjProperty("VALUE.TPALL", "FORMULA", frmcodeingred(x), allergen, "PARAM_CODE")
                        if (allergenValue = "7" or allergenValue = "6") then
                        
                            Dim allergenppmX as Object = ObjProperty("ATTRIBUTE17.TPALL", "FORMULA", frmcodeingred(x), allergen, "PARAM_CODE")
                            If IsBlank(allergenppmX) = 0 Then
                                
                                Dim multipAllergenppmX As Double = Math.Round((CDbl(itmQty(x)) * CDbl(allergenppmX) / 1000), 8, MidpointRounding.AwayFromZero)
                                'MessageList("Multiplication of quantity and " & allergen & " of formula ppmX is: " & multipAllergenppmX)
                                
                                If allergenppmxvalues.ContainsKey(allergen) Then
                                    allergenppmxvalues(allergen) += multipAllergenppmX
                                Else
                                    allergenppmxvalues.Add(allergen, multipAllergenppmX)
                                End If
                                
                            End If
                        End If
                        
                    End If
                Next
                
            Next
        end if
            
        For Each allergen As String In allergenppmxvalues.Keys
            
            Dim totalPpmMax As Double = allergenppmxvalues(allergen)
            ObjPropertySet(totalPpmMax , 0 , "ATTRIBUTE17.TPALL" , "FORMULA" , _OBJECTKEY , allergen , "PARAM_CODE" )
            
        Next
        
        '===========================================================================
        'End: Calculation for Formula ppm MAX  
        '===========================================================================
        
    return 111
    catch ex as exception
        messagelist("Error in calculation: "& ObjProperty("FORMULA_CODE") &"; " & ex.message)
        return 9111
    End Try


End Function


End Class
'========================================================================================================================================================
'End : Reza - Second Workflow - This is the automated Workflow for massive calculation of the Allergen Parameters of all formulas in the table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================