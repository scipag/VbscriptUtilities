wscript.echo Download("https://192.168.0.1/", "", "", "test.html")

Function Download(sUrl, sUser, sPass, sFilename)
	Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

	objXMLHTTP.Open "GET", sUrl, false, sUser, sPass
	objXMLHTTP.Send()
	Do Until objXMLHTTP.Status = 200 : Wscript.Sleep(1000) : Loop

	If objXMLHTTP.Status = 200 Then
		Set objADOStream = CreateObject("ADODB.Stream")
		objADOStream.Open
		objADOStream.Type = 1
		objADOStream.Write objXMLHTTP.ResponseBody
		objADOStream.Position = 0    
 
	    Set objFSO = Createobject("Scripting.FileSystemObject")
		If objFSO.Fileexists(sFilename) Then objFSO.DeleteFile sFilename
		Set objFSO = Nothing
 
		objADOStream.SaveToFile sFilename
		objADOStream.Close 
		Set objADOStream = Nothing
	End if

	Set objXMLHTTP = Nothing
End Function