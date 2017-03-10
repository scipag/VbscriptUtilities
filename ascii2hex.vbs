Wscript.Echo hexEncode("This is a simple test!")

Function hexEncode(str)
	Dim strEncoded, i
	strEncoded = "0x"

	For i = 1 to Len(str)
		If (i Mod 4) = 1 Then
			strEncoded = strEncoded + " 0x"
		End If	

		strEncoded = strEncoded + Hex(Asc(Mid(str, i, 1)))
	Next

	hexEncode = strEncoded
End Function
