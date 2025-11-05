
'Automating the Rounding/Truncating according to the ingredients classifications in the save Script of the Formula


'===========================================================================
'Begin: Reza - 4 Cifre Logic for BOMs Quantities  
'===========================================================================
'messagelist("3")
Dim cifreCheck As Boolean = False
Dim itmuom as Object = objproperty("UOMCODE.INGR", "", "", "*", 1)

For x As Integer = 0 To itmuom.Length - 1
   'Dim itmuom As String = objproperty("UOMCODE", "ITEM", ingritemCodes(x))
	Dim rawQty As String = itmQty(x).ToString().Replace(",", ".")
	Dim qty As Double = Convert.ToDouble(rawQty, Globalization.CultureInfo.InvariantCulture)
	'messagelist("itmuom: " & itmuom(x))

	If itmuom(x) = "EA" Or  itmuom(x) = "PCS" Then

	   ' messagelist("qty: " & qty.ToString(Globalization.CultureInfo.InvariantCulture))
		
		' Check if more than 4 decimals and truncate it
		Dim truncatedvalue As Double = Math.Truncate(qty * 1000) / 1000
		If truncatedvalue <> qty Then      
			
			'messagelist("Truncated qty: " & truncatedvalue.ToString(Globalization.CultureInfo.InvariantCulture))

			ObjPropertySet(1, 0, "FOUR_DECIMAL_INGRD", "ITEM", ingritemCodes(x))
			ObjPropertySet(truncatedvalue, 0, "QUANTITY.INGR.A", "", "",ingritemCodes(x),"ITEMCODE")

			cifreCheck = True
			messagelist("The Quantity for ingredient " & ingritemCodes(x) & " was " & qty & " and Truncated to " & truncatedvalue)
			messagelist("The 4Cifre flag for item " + ingritemCodes(x) + " enabled")
		End If
		
	Else

		'messagelist("qty: " & qty.ToString(Globalization.CultureInfo.InvariantCulture))

		' Check if more than 4 decimals
		If Math.Round(qty, 3) <> qty Then      
			' Round to 3 decimals
			Dim roundedQty As Double = Math.Round(qty, 3, MidpointRounding.AwayFromZero)

			'messagelist("rounded qty: " & roundedQty.ToString(Globalization.CultureInfo.InvariantCulture))

			ObjPropertySet(1, 0, "FOUR_DECIMAL_INGRD", "ITEM", ingritemCodes(x))
			ObjPropertySet(roundedQty, 0, "QUANTITY.INGR.A", "", "",ingritemCodes(x),"ITEMCODE")

			cifreCheck = True
			messagelist("The Quantity for ingredient " & ingritemCodes(x) & " was " & qty & " and Rounded to " & roundedQty)
			messagelist("The 4Cifre flag for item " + ingritemCodes(x) + " enabled")
		End If
	End If
Next

If cifreCheck Then
   ObjPropertySet(1, 0, "FOUR_DECIMAL_INGRD", "", "")
End If

'===========================================================================
'End: Reza - 4 Cifre Logic for BOMs Quantities   
'===========================================================================