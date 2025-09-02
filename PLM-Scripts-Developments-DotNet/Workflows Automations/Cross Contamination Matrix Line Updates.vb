'CROSS CONTAMINATION MATRIX Updates

Option Strict Off
imports System
imports System.Diagnostics
imports Microsoft.Visualbasic
imports System.Data
Imports System.Collections.Generic
Imports System.Data.Common
Imports System.Threading
Imports System.Linq


Class ActionScript
 Inherits FcProcFuncSetEventWF

Function wf_start() As Long
    
Try

    '=============================
    'Updating From FG Workcenters
    '=============================
    Dim oLoc As Object = ObjProperty("ATTRIBVAL.CONTEXT", "", "", "*", 1)
    Dim oLines As Object = ObjProperty("FIELD1.MATRIX.V\DMFORMULA5", "FORMULA", "", "*")
    Dim dtTemplate As DataTable = TableLookupEx("SO_CROSS_CONT_MAT_IR_IT", "TableLookup", "@TFORMULA\001")
    Dim CCsiteCodes As New HashSet(Of String)
    Dim removeSiteList as New HashSet(Of String) 'this is the list to be checked finaly and remove the extra CC
    
    if oLines.Length > 0 Then 'remove the whole lines to update if there were already some rows in CC
         RemoveCustomTable("C_CROSS_CONTAMINATION_MATRIX_IR_IT")
    end if
    
    For i As Integer = 0 To oLoc.Length - 1   'Loop over each workcenter
        
        For Each temRow As DataRow In dtTemplate.Rows  'Loop over rows of the template CC list to find related line to the WC
            
            If InStr(1, temRow("SITE_CODE").ToString(), oLoc(i), vbTextCompare) > 0 Then
                Dim newDM0Row As DataRow = GetNewRow("", "", "MATRIX.V\DMFORMULA5")
                newDM0Row("FIELD1") = temRow("SITE_CODE")
                newDM0Row("FIELD2") = temRow("ALLERGEN_SOY")
                newDM0Row("FIELD3") = temRow("ALLERGEN_EGG")
                newDM0Row("FIELD4") = temRow("ALLERGEN_MILK") 
                newDM0Row("FIELD5") = temRow("ALLERGEN_CELERY")
                newDM0Row("FIELD6") = temRow("ALLERGEN_CRUSTACEAEN_")
                newDM0Row("FIELD7") = temRow("ALLERGEN_LUPIN")
                newDM0Row("FIELD8") = temRow("ALLERGEN_FISH")
                newDM0Row("FIELD9") = temRow("ALLERGEN_MOLLUSCS")
                newDM0Row("FIELD10") = temRow("ALLERGEN_MUSTARD")
                newDM0Row("FIELD11") = temRow("ALLERGEN_PEANUTS")
                newDM0Row("FIELD12") = temRow("ALLERGEN_SESAME")
                newDM0Row("FIELD13") = temRow("ALLERGEN_SO2")
                newDM0Row("FIELD14") = temRow("ALLERGEN_TREENUTS")
                newDM0Row("FIELD15") = temRow("ALLERGEN_WHEAT")
                newDM0Row("FIELD16") = temRow("ALLERGEN_ASPARTAME")
                'messagelist("1: ",newDM0Row("FIELD1"))
                CommitNewRow("", "", "MATRIX.V\DMFORMULA5", newDM0Row)
                'messagelist("2: ",temRow("SITE_CODE"))
            End If
        Next
    Next
    
    
    Dim updatedCCLines As Object = ObjProperty("FIELD1.MATRIX.V\DMFORMULA5", "FORMULA", "", "*")
    for x as integer = 0 to updatedCCLines.Length -1 'check this in order to avoid adding new line which is already exists
        CCsiteCodes.Add(updatedCCLines(x))
    Next
    
    '====================================
    'Updating From SF's CC in the FG Boms
    '====================================
    Dim fgComponentBoms as Object = ObjProperty("COMPONENTIND.INGR", "", "", "*", 1)
    Dim fgFormulasBoms as Object = ObjProperty("FORMULACODE.INGR", "", "", "*", 1)
    
    For r as integer = 0 to fgComponentBoms.Length -1
        if fgComponentBoms(r) = "1" Then
            Dim dtSFCc As DataTable = TableLookupEx("SO_CROSS_CONT_MAT", "TableLookup", fgFormulasBoms(r))
            if dtSFCc.Rows.Count > 0 then 
                For Each rowSFCc as DataRow in dtSFCc.Rows
                    
                    If Not CCsiteCodes.Contains(rowSFCc("SITE_CODE")) Then  'check if the line is already in the FG CC lines
                        'messagelist("RowSf: ",rowSFCc("SITE_CODE"))
                        Dim newDM0Row As DataRow = GetNewRow("", "", "MATRIX.V\DMFORMULA5")
                        newDM0Row("FIELD1") = rowSFCc("SITE_CODE")
                        newDM0Row("FIELD2") = rowSFCc("ALLERGEN_SOY")
                        newDM0Row("FIELD3") = rowSFCc("ALLERGEN_EGG")
                        newDM0Row("FIELD4") = rowSFCc("ALLERGEN_MILK") 
                        newDM0Row("FIELD5") = rowSFCc("ALLERGEN_CELERY")
                        newDM0Row("FIELD6") = rowSFCc("ALLERGEN_CRUSTACEAEN_")
                        newDM0Row("FIELD7") = rowSFCc("ALLERGEN_LUPIN")
                        newDM0Row("FIELD8") = rowSFCc("ALLERGEN_FISH")
                        newDM0Row("FIELD9") = rowSFCc("ALLERGEN_MOLLUSCS")
                        newDM0Row("FIELD10") = rowSFCc("ALLERGEN_MUSTARD")
                        newDM0Row("FIELD11") = rowSFCc("ALLERGEN_PEANUTS")
                        newDM0Row("FIELD12") = rowSFCc("ALLERGEN_SESAME")
                        newDM0Row("FIELD13") = rowSFCc("ALLERGEN_SO2")
                        newDM0Row("FIELD14") = rowSFCc("ALLERGEN_TREENUTS")
                        newDM0Row("FIELD15") = rowSFCc("ALLERGEN_WHEAT")
                        newDM0Row("FIELD16") = rowSFCc("ALLERGEN_ASPARTAME")
                        CommitNewRow("", "", "MATRIX.V\DMFORMULA5", newDM0Row)
                    end if
                Next
            else 'Notice the user if in the SF CC there is no line
                messagelist("There is no Cross Contamination Line in the lower level Formuala " & fgFormulasBoms(r))
            end if
        end if 
    Next
    messagelist("The Cross Contamination Lines has been Updated")
    
    return 111

Catch ex as exception
    messagelist("Error: "+ ex.message)
    PublishDebuggingInfo("Error: " + ex.message)
    Return 1
End Try
    

End Function


End Class