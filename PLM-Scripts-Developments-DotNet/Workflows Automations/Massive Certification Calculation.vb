'Description: Massive workflow for automatic calculation of Certification parameters of all formulas 
'Note: first the first workflow is launching and then for each formula in the defined table the second workflow will be launched automatically

'========================================================================================================================================================
'Start : Reza - first workflow - This is the automated Workflow for massive calculation of the Certification Parameters of all formulas in the Custom table 'SO_MASSIVE_ALLERGENS_CALC'
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
'End : Reza - first workflow - This is the automated Workflow for massive calculation of the Certification Parameters of all formulas in the Custom table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================

'========================================================================================================================================================
'Begin : Reza - Second Workflow - This is the automated Workflow for massive calculation of the Certification Parameters of all formulas in the Custom table 'SO_MASSIVE_ALLERGENS_CALC'
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
    Try
        Dim ingritemsCodes As Object = objproperty("ITEMCODE.INGR", "", "", "*", 1)
    
        Dim certificationParams As New List(Of String) From {"HALAL", "HALAL_PACK_LOGO", "KOSHER", "KOSHER_PASSOVER", "KOSHER_PACK_LOGO", "BIO", "FAIRTRADE", "RSPO_MASSBALANCE", "VEGAN", "VEGETARIAN", "UTZ", "IP", "IGP", "DOP", "VLOG", "AIC"}
        
        Dim certificationValues As New Dictionary(Of String, List(Of Integer)) ' Store the values for each certification

        If ingritemsCodes.Length > 0 Then
            For x As Integer = 0 To ingritemsCodes.Length - 1
                For Each certification As String In certificationParams
                    Dim certifValue As Object = ObjProperty("VALUE.TPALL", "ITEM", ingritemsCodes(x), certification, "PARAM_CODE")
                    Dim value As Integer = If(IsNumeric(certifValue), CInt(certifValue), 0)
                    
                    ' Ensure the certification key exists in the dictionary before adding values
                    If Not certificationValues.ContainsKey(certification) Then
                        certificationValues(certification) = New List(Of Integer)()
                    End If
                    
                    ' Add the value to the list for the current certification
                    certificationValues(certification).Add(value)
                Next
            Next

            For Each certification As String In certificationParams
                ' Safely get the list of values for the current certification
                Dim values As List(Of Integer) = Nothing
                If certificationValues.TryGetValue(certification, values) Then
				
                    ' Set the property with the minimum value from the list
                    ' If the values list is not empty, find the minimum value
                    If values.Count > 0 Then
                        Dim minValue As Integer = values.Min()
                        ObjPropertySet(minValue, 0, "VALUE.TPALL", "", "", certification, "PARAM_CODE")
                    End If
					
                End If
            Next

            PublishDebuggingInfo((objproperty("FORMULA_CODE")) & " has calculated")
        Else
            PublishDebuggingInfo(objproperty("FORMULA_CODE") & " does NOT have ingredients")
        End If

        Return 111
    Catch ex As Exception
        MessageList("Error in calculation: " & ObjProperty("FORMULA_CODE") & "; " & ex.Message)
        Return 9111
    End Try
End Function

End Class

'========================================================================================================================================================
'Begin : Reza - Second Workflow - This is the automated Workflow for massive calculation of the Allergen Parameters of all formulas in the table 'SO_MASSIVE_ALLERGENS_CALC'
'========================================================================================================================================================