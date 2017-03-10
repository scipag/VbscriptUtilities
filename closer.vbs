Set objWmi = GetObject("winmgmts:")

process = Inputbox("Which application shall be closed?", "Application to Close", "spotify.exe")
delay   = Inputbox("In how many minutes shall the closure be?", "Delay to Close", "10")

WScript.Sleep (delay * 1000 * 60)

strWmiq = "select * from Win32_Process where name='" & process & "'"
Set objQResult = objWmi.Execquery(strWmiq)

For Each objProcess In objQResult
	Ret = objProcess.Terminate(1)
Next
