Wscript.Echo "VBShell 1.2"
Wscript.Echo "Copyright (c) 2012-2013 Marc Ruef"

Do While True
	'Read input line
	Wscript.Stdout.Write(">>>")
	ln = Wscript.Stdin.Readline

	If Lcase(Trim(ln)) = "exit" OR Lcase(Trim(ln)) = "quit" Then
		Exit Do
	ElseIf Lcase(Trim(ln)) = "help" Then
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo "Buit-In Commands"
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo "exit" & vbTab & vbTab & vbTab & "quit shell"
		Wscript.Echo
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo "Extended Functions"
		Wscript.Echo "---------------------------------------------------"
		Wscript.Echo "hexencode(sString)" & vbTab & "convert string to hex"
		Wscript.Echo "md5(sString)" & vbTab & vbTab & "generate md5 hash"
		Wscript.Echo "processkill(sProcess)" & vbTab & "terminate a process by name"
		Wscript.Echo "readfile(sFile)" & vbTab & vbTab & "read file to string"
		Wscript.Echo "sha1(sString)" & vbTab & vbTab & "generate sha1 hash"
		Wscript.Echo "sleep(iSeconds)" & vbTab & vbTab & "wait for a few seconds"
	ElseIf LenB(ln) Then
		On Error Resume Next
		Err.Clear

		'Execute line
		If InStr(2, ln, " = ") OR InStr(2, ln, Lcase(" then")) Then
			Execute ln
		Else
			Wscript.Echo Eval(ln)
		End If
		
		If Err.Number <> 0 Then
			Wscript.Echo("Error Code #" & Err.Number & ": " & Err.Description & " [" & Mid(ln, 1, 20) & "]")
			On Error Goto 0
		End If
	End If
Loop

'''Remapping existing functions

Function repeat(i, s)
	repeat = String(i, s)
End Function


'''Extended Functions

Function readfile(sFile)
	Set objFileSystem = CreateObject("Scripting.FileSystemObject")
	Set objInputFile = objFileSystem.OpenTextFile(sFile, 1)
	readfile = objInputFile.ReadAll
	objInputFile.Close
	Set objFileSystem = Nothing
End Function

Function hexencode(sAscii)
	For i = 1 to Len(sAscii)
		If (i Mod 4) = 1 Then
			strEncoded = Trim(strEncoded & " 0x")
		End If
		strEncoded = strEncoded & Hex(Asc(Mid(sAscii, i, 1)))
	Next
	hexencode = strEncoded
End Function

Function processkill(sProcess)
	Set objShell = CreateObject("WScript.Shell")
	Set objWmi = GetObject("winmgmts:")
	strWmiq = "select * from Win32_Process where name='" & sProcess & "'"
	Set objQResult = objWmi.Execquery(strWmiq)
	For Each objProcess In objQResult
		processkill = objProcess.Terminate(1)
	Next
End Function

Function md5(sString) 
	Dim asc, enc, bytes, s, pos 
	Set asc = CreateObject("System.Text.UTF8Encoding") 
	Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider") 
	bytes = asc.GetBytes_4(sString) 
	bytes = enc.ComputeHash_2((bytes)) 
	s = ""
	For pos = 1 To Lenb(bytes) 
		s = s & LCase(Right("0" & Hex(Ascb(Midb(bytes, pos, 1))), 2)) 
	Next 
	Set asc = Nothing 
	Set enc = Nothing 
	md5 = s 
End Function

Function sha1(sString) 
	Dim asc, enc, bytes, s, pos 
	Set asc = CreateObject("System.Text.UTF8Encoding") 
	Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider") 
	bytes = asc.GetBytes_4(sString) 
	bytes = enc.ComputeHash_2((bytes)) 
	s = ""
	For pos = 1 To Lenb(bytes) 
		s = s & LCase(Right("0" & Hex(Ascb(Midb(bytes, pos, 1))), 2)) 
	Next 
	Set asc = Nothing 
	Set enc = Nothing 
	sha1 = s 
End Function

Function increment(sStart, sEnd)
	For i = sStart To sEnd
		s = s & i & vbCrLf
	Next

	increment = s
End Function

Function sleep(iSeconds)
	WScript.Sleep (iSeconds * 1000)
End Function
